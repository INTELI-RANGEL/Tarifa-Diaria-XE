unit frm_oficiosdemovimientos_detalles;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, DBGrids, frm_connection, StdCtrls, DBCtrls,
  ExtCtrls, DB, Global, Menus, 
  ZAbstractRODataset, ZAbstractDataset, ZDataset, UdbGrid,
  unittbotonespermisos, 
  DBDateTimePicker;

type
  TfrmOficiosDeMovimientos_detalles = class(TForm)
    grid_plataformas: TDBGrid;
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
    ds_gruposdepersonal: TDataSource;
    zq_gruposdepersonal: TZQuery;
    zq_gruposdepersonaliId: TIntegerField;
    zq_gruposdepersonalsContrato: TStringField;
    zq_gruposdepersonalsNombreOficio: TStringField;
    zq_gruposdepersonalsOficioDeAutorizacion: TStringField;
    zq_gruposdepersonaldFecha: TDateField;
    zq_gruposdepersonaldFechaInicial: TDateField;
    zq_gruposdepersonaldFechaFinal: TDateField;
    zq_gruposdepersonalsDescripcion: TStringField;
    zq_gruposdepersonaleTipo: TStringField;
    dBPersonal: TDBGrid;
    rbC14: TRadioButton;
    rbc15: TRadioButton;
    rbcPernocta: TRadioButton;
    lbl_recursos: TLabel;
    ds_MovtosPersonal: TDataSource;
    movtosPersEq: TZQuery;
    movtosPersEqdFechaVigencia: TDateField;
    movtosPersEqsAnexo: TStringField;
    movtosPersEqsNumeroActividad: TStringField;
    movtosPersEqiItemOrden: TIntegerField;
    movtosPersEqsContrato: TStringField;
    movtosPersEqsNumeroOrden: TStringField;
    movtosPersEqsDescripcion: TStringField;
    movtosPersEqdCantidad: TFloatField;
    movtosPersEqAnterior: TFloatField;
    movtosPersEqiFolioOficio: TIntegerField;
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure grid_plataformasMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure grid_plataformasMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure grid_plataformasTitleClick(Column: TColumn);
    procedure Copy1Click(Sender: TObject);
    procedure Paste1Click(Sender: TObject);
  private
  sMenuP: String;
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmOficiosDeMovimientos_detalles: TfrmOficiosDeMovimientos_detalles;
   UtGrid:TicDbGrid;
   botonpermiso: tbotonespermisos;
   sOpcion, lStatusOrig: string;
implementation

{$R *.dfm}

procedure TfrmOficiosDeMovimientos_detalles.FormShow(Sender: TObject);
begin
  sMenuP:=stMenu;
  BotonPermiso := TBotonesPermisos.Create(Self, connection.zConnection, global_grupo, 'cGruposdePersonal', PopupPrincipal);
  OpcButton := '' ;

  zq_gruposdepersonal.active := false ;
  zq_gruposdepersonal.Open;
  UtGrid:=TicdbGrid.create(grid_PLATAFORMAS);
//  BotonPermiso.permisosBotones(frmBarra1);
end;

procedure TfrmOficiosDeMovimientos_detalles.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  zq_gruposdepersonal.Cancel ;
  action := cafree ;
  utgrid.Destroy;
  botonpermiso.Free;
end;

procedure TfrmOficiosDeMovimientos_detalles.grid_plataformasMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
    UtGrid.dbGridMouseMoveCoord(x,y);
end;

procedure TfrmOficiosDeMovimientos_detalles.grid_plataformasMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  UtGrid.DbGridMouseUp(Sender,Button,Shift,X, Y);
end;

procedure TfrmOficiosDeMovimientos_detalles.grid_plataformasTitleClick(Column: TColumn);
begin
  UtGrid.DbGridTitleClick(Column);
end;

procedure TfrmOficiosDeMovimientos_detalles.Paste1Click(Sender: TObject);
begin
   UtGrid.AddRowsFromClip;
end;

procedure TfrmOficiosDeMovimientos_detalles.Copy1Click(Sender: TObject);
begin
    UtGrid.CopyRowsToClip;
end;

end.

