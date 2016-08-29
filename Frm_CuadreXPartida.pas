unit Frm_CuadreXPartida;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, NxColumns, NxColumnClasses, Grids, AdvGrid,
  NxScrollControl, NxCustomGridControl, NxGrid, ExtCtrls, DB,
  ZAbstractRODataset, ZDataset, StdCtrls, DBCtrls, ComCtrls, DBDateTimePicker,
  JvExStdCtrls, NxPageControl, Menus, AdvOfficePager,
  JvDBCombobox, AdvGlowButton, ImgList, NxCollection,
  Newpanel, AdvStickyPopupMenu, DBGrids,
  JvExDBGrids, JvDBGrid, JvDBUltimGrid, Buttons, NxEdit,UTSuperPanel,ComObj,ShlObj,
  DateUtils, JvCombobox;

type
  TFrmCuadreXPartida = class(TForm)
    pnl1: TPanel;
    ds_ordenesdetrabajo: TDataSource;
    ordenesdetrabajo: TZReadOnlyQuery;
    Label1: TLabel;
    tdIdFecha: TDBDateTimePicker;
    Label2: TLabel;
    tsIdCategorias: TComboBox;
    pmPage: TPopupMenu;
    AgregarNuevaPernocta1: TMenuItem;
    pmEncabezados: TPopupMenu;
    OcultarColumna1: TMenuItem;
    dbcmbPernoctas: TJvDBComboBox;
    lblPernocta: TLabel;
    pnl2: TPanel;
    pnl3: TPanel;
    NxPgcDatos: TNxPageControl;
    tmrRecargar: TTimer;
    ImgBtns: TImageList;
    btnPost: TAdvGlowButton;
    btnDelete: TAdvGlowButton;
    btnPernocta: TNxButton;
    AvSPpmPernoctas: TAdvStickyPopupMenu;
    EliminarCategoria1: TMenuItem;
    N1: TMenuItem;
    ImportarCategoria1: TMenuItem;
    DeUndiaAnterior1: TMenuItem;
    ElegirDia1: TMenuItem;
    N2: TMenuItem;
    InsertarValoresenBlanco1: TMenuItem;
    NGbxConsulta: tNewGroupBox;
    pnl4: TPanel;
    dbUgridBuscar: TJvDBUltimGrid;
    pnl5: TPanel;
    cmdCancelar: TButton;
    cmdAceptar: TButton;
    btnNuevoPersonal: TBitBtn;
    catalogo_maestro: TZReadOnlyQuery;
    dsMasterCatalog: TDataSource;
    NxBtnEdtBuscar: TNxButtonEdit;
    lbl2: TLabel;
    btnExportar: TButton;
    SdgExcel: TSaveDialog;
    btnImportar: TButton;
    dlgOpenXls: TOpenDialog;
    procedure FormShow(Sender: TObject);
    procedure tdIdFechaExit(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure tmrRecargarTimer(Sender: TObject);
    procedure btnPostClick(Sender: TObject);
    procedure btnPernoctaClick(Sender: TObject);
    procedure btnPernoctaMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure AvSPpmPernoctasButtonBar0Click(Sender: TObject);
    procedure AvSPpmPernoctasMenuHide(Sender: TObject);
    procedure NxPgcDatosChange(Sender: TObject);
    procedure AgregarNuevaPernocta1Click(Sender: TObject);
    procedure cmdCancelarClick(Sender: TObject);
    procedure NxBtnEdtBuscarButtonClick(Sender: TObject);
    procedure cmdAceptarClick(Sender: TObject);
    procedure EliminarCategoria1Click(Sender: TObject);
    procedure DeUndiaAnterior1Click(Sender: TObject);
    procedure InsertarValoresenBlanco1Click(Sender: TObject);
    procedure btnDeleteClick(Sender: TObject);
    procedure NxPgcDatosPageClose(Sender: TObject; PageIndex: Integer;
      var Accept: Boolean);
    procedure tsIdCategoriasExit(Sender: TObject);
    procedure tsNumeroOrdenExit(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormDestroy(Sender: TObject);
    procedure btnExportarClick(Sender: TObject);
    procedure SdgExcelTypeChange(Sender: TObject);
    procedure btnImportarClick(Sender: TObject);
  private
    { Private declarations }
    procedure CargaDatos(sParamContrato,sParamOrden:string;sParamDatos:TsuperPanel);
    procedure InicarCategoria;
    procedure RecarGarDatos;
    procedure GrabarInformacion;
    Function ValidarCategoria(var Excel:Variant; var Hoja:Variant;sTipo:string;dParamFecha:TDate):Boolean;
  public
    { Public declarations }
    sFrente:string;
    dFecha:TDate;
  end;


var
  FrmCuadreXPartida: TFrmCuadreXPartida;
  ArrayTipo: array[1..3, 1..5] of string;
  iTipo : integer;
  //sIdPernocta:string;
  {ListaPdas,}
  ListaObj:TStringList;
  bRecargar:Boolean;
  isChange,Cerrando:Boolean;
  btnx,btny:integer;
  MnIClick,Recharged:Boolean;
  OldFecha:TDate;
  OldCategoria:string;
  OldFrente:string;
  sTipoFiltro:string;
 // iAuxPos,RowAnt:Integer;

implementation

uses global, frm_connection, UnitExcel, masUtilerias;

{$R *.dfm}

procedure TFrmCuadreXPartida.GrabarInformacion;
var
  i,Reng,iColPos:Integer;
  SpPanelDatos:TSuperPanel;
  cadenamensajes,sUltimPer,cadenamensaje : string;
  pPartida:TPartida;
  CatCero,AlMenos1:Boolean;
  QRecurso:TZQuery;
begin
  
  QRecurso:=TZQuery.Create(nil);
  QRecurso.Connection:=connection.zConnection;

  //Verificamos si el reporte de barco está validado...
  connection.zCommand.Active := False;
  connection.zCommand.SQL.Clear;
  connection.zCommand.SQL.Add('select * from reportediario where sContrato =:Contrato and dIdFecha =:fecha and lStatus ="Autorizado"');
  connection.zCommand.ParamByName('Contrato').AsString := global_Contrato_Barco;
  connection.zCommand.ParamByName('fecha').AsDate      := tdIdFecha.Date;
  connection.zCommand.Open;

  if connection.zCommand.RecordCount = 0 then
  begin
   //if not Cerrando then
      //if Messagedlg('¿Desea Conservar las Categorias que Tienen Cero Reportado?',mtConfirmation,[mbYes,mbNo],0)=Mryes then
    CatCero:=false;
        
    for I := 0 to ListaObj.Count-1 do
    begin
      SpPanelDatos:=TSuperPanel(ListaObj.Objects[i]);
      //sUltimPer:=SpPanelDatos.Id;
      sUltimPer:='';
      //Consultamos la Vigencia de Barco
      connection.QryBusca.Active := False;
      connection.QryBusca.SQL.Clear;
      connection.QryBusca.sql.text:='select * from embarcacion_vigencia ' +
                                    'where scontrato=:Contrato and dFechaInicio='+
                                    '(select max(dFechaInicio) from embarcacion_vigencia ' +
                                    'where scontrato=:Contrato and dFechaInicio<=:Fecha)';
      connection.QryBusca.ParamByName('contrato').AsString:=global_Contrato_Barco;
      connection.QryBusca.ParamByName('Fecha').AsDate:=tdIdFecha.Date;
      connection.QryBusca.Open;

      if connection.QryBusca.RecordCount=1 then
        sUltimPer:=connection.QryBusca.FieldByName('sIdEmbarcacion').AsString;


      connection.QryBusca.ParamByName('Contrato').AsString := SpPanelDatos.sContrato;
      connection.QryBusca.ParamByName('orden').AsString    := SpPanelDatos.sNumeroOrden;
      connection.QryBusca.ParamByName('fecha').AsDate      := tdIdFecha.Date;
      connection.QryBusca.ParamByName('turno').AsString    := global_turno;
      connection.QryBusca.Open;



      //Consultamos el iIdDiario de las actividades para insertar los datos de acuerdo al horario..
      connection.QryBusca.Active := False;
      connection.QryBusca.SQL.Clear;
      connection.QryBusca.SQL.Add('select iIdDiario from bitacoradeactividades where sContrato =:Contrato and dIdFecha =:Fecha '+
                                  'and sNumeroOrden =:Orden and sIdTurno =:Turno and sIdTipoMovimiento = "'+ArrayTipo[iTipo,5]{"E","N"}+'" ');
      connection.QryBusca.ParamByName('Contrato').AsString := SpPanelDatos.sContrato;
      connection.QryBusca.ParamByName('orden').AsString    := SpPanelDatos.sNumeroOrden;
      connection.QryBusca.ParamByName('fecha').AsDate      := tdIdFecha.Date;
      connection.QryBusca.ParamByName('turno').AsString    := global_turno;
      connection.QryBusca.Open;

      if connection.QryBusca.RecordCount > 0 then
      begin
        //Primero eliminamos el personal y equipo de la tabla de bitacoradepersonal..
        Connection.qryBusca2.SQL.Clear;
        Connection.qryBusca2.SQL.Add( 'Select '+ArrayTipo[iTipo,3]{sIdpersonal,sIdequipo}+' from '+ArrayTipo[iTipo,2]{personal,equipo}+
                                      ' Where sContrato =:Contrato and lCobro ="'+ArrayTipo[iTipo,4]{"Si","No"}+'" ');
        Connection.qryBusca2.Params.ParamByName('Contrato').Datatype := ftString;
        Connection.qryBusca2.Params.ParamByName('Contrato').Value    := SpPanelDatos.sContrato;
        Connection.qryBusca2.Open;

        while not connection.QryBusca2.Eof do
        begin
            connection.zCommand.Active := False;
            connection.zCommand.SQL.Clear;
            connection.zCommand.SQL.Add('Delete from '+ArrayTipo[iTipo,1]{bitacoradepersonal,bitacoradeequipos}+' '+
                        'where sContrato =:contrato and dIdFecha =:fecha and sidPernocta=:pernocta and '+ArrayTipo[iTipo,3]{sIdpersonal,sIdequipo}+' =:Personal ');
            connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
            connection.zCommand.Params.ParamByName('Contrato').Value    := SpPanelDatos.sContrato;
            connection.zCommand.Params.ParamByName('Fecha').DataType    := ftDate;
            connection.zCommand.Params.ParamByName('Fecha').Value       := tdIdFecha.Date;
            connection.zCommand.Params.ParamByName('Personal').DataType := ftString;
            connection.zCommand.Params.ParamByName('Personal').Value    := Connection.qryBusca2.FieldValues[ArrayTipo[iTipo,3]{sIdpersonal,sIdequipo}];
            connection.zCommand.Params.ParamByName('Pernocta').AsString    := sUltimPer;//LoboAzul27
            connection.zCommand.ExecSQL;
            connection.QryBusca2.Next;
        end;

        Reng:=0;
        while Reng<SpPanelDatos.NxGridDatos.RowCount do
        begin
          for iColPos := 2 to SpPanelDatos.NxGridDatos.Columns.Count - 1 do
          begin
            with SpPanelDatos.NxGridDatos do
            begin

              if StrToFloatDef(Cell[iColPos,Reng].AsString,0)<>0 then
              begin
                AlMenos1:=True;
                QRecurso.Active := False;
                QRecurso.SQL.Clear;
                pPartida:=TPartida(SpPanelDatos.ListaPdas.Objects[SpPanelDatos.ListaPdas.IndexOf(IntToStr(iColPos))]);

                if iTipo <> 2 then
                   QRecurso.SQL.Add('insert into '+ArrayTipo[iTipo,1]{bitacoradepersonal,bitacoradeequipo}+' (sContrato, dIdFecha, iIdDiario, sIdPersonal, iItemOrden, sDescripcion, sIdPernocta, sIdPlataforma, sHoraInicio, sHoraFinal, dCantidad, sAgrupaPersonal) '+
                                     'values (:Contrato, :fecha, :diario, :personal, :item, :descripcion, :pernocta, :plataforma, :inicio, :final, :Cantidad, :agrupa)')
                else
                   QRecurso.SQL.Add('insert into '+ArrayTipo[iTipo,1]{bitacoradepersonal,bitacoradeequipo}+' (sContrato, dIdFecha, iIdDiario, sIdEquipo, iItemOrden, sDescripcion, sIdPernocta, sHoraInicio, sHoraFinal, dCantidad) '+
                                     'values (:Contrato, :fecha, :diario, :personal, :item, :descripcion, :pernocta, :inicio, :final, :Cantidad)');
                QRecurso.ParamByName('Contrato').AsString    := SpPanelDatos.scontrato;//global_contrato;
                QRecurso.ParamByName('fecha').AsDate         := tdIdFecha.Date;
                QRecurso.ParamByName('diario').AsInteger     := pPartida.IdDiario;  //rxPersonal.FieldValues['iIdDiario'+IntToStr(indice)];
                QRecurso.ParamByName('personal').AsString    := TRecurso(SpPanelDatos.ListaRecursos.Objects[SpPanelDatos.ListaRecursos.IndexOf(cell[0,Reng].AsString)]).sIdRecurso;//cell[0,Reng].AsString;  //rxPersonal.FieldValues[ArrayTipo[iTipo,3]{sIdPersonal,sIdEquipo}];
                QRecurso.ParamByName('item').AsInteger       := TRecurso(SpPanelDatos.ListaRecursos.Objects[SpPanelDatos.ListaRecursos.IndexOf(cell[0,Reng].AsString)]).ItemOrden;
                QRecurso.ParamByName('descripcion').AsString := TRecurso(SpPanelDatos.ListaRecursos.Objects[SpPanelDatos.ListaRecursos.IndexOf(cell[0,Reng].AsString)]).sDescripcion;//  rxPersonal.FieldValues['sDescripcion'];
                QRecurso.ParamByName('pernocta').AsString    := sUltimPer;//rxPersonal.FieldValues['sIdPernocta'];
                if iTipo <> 2 then
                begin
                    QRecurso.ParamByName('plataforma').AsString  := TRecurso(SpPanelDatos.ListaRecursos.Objects[SpPanelDatos.ListaRecursos.IndexOf(cell[0,Reng].AsString)]).sPlataforma;
                    QRecurso.ParamByName('agrupa').AsString      := TRecurso(SpPanelDatos.ListaRecursos.Objects[SpPanelDatos.ListaRecursos.IndexOf(cell[0,Reng].AsString)]).sAgrupa;
                end;
                QRecurso.ParamByName('inicio').AsString      := pPartida.sHoraInicio;   //rxPersonal.FieldValues['sHoraInicio'+IntToStr(indice)];
                QRecurso.ParamByName('final').AsString       := pPartida.sHoraFinal;   //rxPersonal.FieldValues['sHoraFinal'+IntToStr(indice)];
                QRecurso.ParamByName('cantidad').AsFloat     := Cell[iColPos,Reng].asfloat;//rxPersonal.FieldValues['dCantidad'+IntToStr(indice)];
                QRecurso.ExecSQL;
              end;
            end;

          end;

          Inc(Reng);
        end;

      end
      else
      begin
          If pos('NO SE ENCONTRARON PARTIDAS REPORTADAS', cadenamensaje) = 0 then
                cadenamensaje := cadenamensaje + 'NO SE ENCONTRARON PARTIDAS REPORTADAS';
      end;
    end;
  end
  else
  begin
      messageDLG('El Reporte Diario está Autorizado!', mtInformation, [mbOk], 0);
      cadenamensaje := 'Autorizado';
  end;

  if cadenamensaje <> 'Autorizado'  then
  begin
    if cadenamensaje = '' then
    begin
      if not Cerrando then
      begin
        //tdIdfecha.OnExit(sender);
        messageDLG('Proceso Terminado con Exito!', mtInformation, [mbOk], 0);
      end;
    end
    else
       messageDLG(cadenamensaje, mtInformation, [mbOk], 0);
  end;

end;

procedure TFrmCuadreXPartida.RecarGarDatos;
var
  QrPernoctas:TZReadOnlyQuery;
  sDatosPanel:TSuperPanel;
  NxTbs: TNxTabSheet;
  StMiPernocta:TStickyMenuItem;
  i:Integer;
  QrFrentes:TZReadOnlyQuery;
  sNombre:string;
begin
  Recharged:=True;
  QrPernoctas:=TZReadOnlyQuery.Create(nil);
  QrFrentes:=TZReadOnlyQuery.Create(nil);
  dbcmbPernoctas.Items.Clear;
  dbcmbPernoctas.Values.Clear;
  with QrFrentes do
  begin
    Connection:=frm_connection.connection.zConnection;
    SQL.Text:='select ot.* from ordenesdetrabajo ot ' +
              'inner join contratos c on(ot.sContrato=c.sContrato) ' +
              'inner join bitacoradeactividades ba on(ba.sContrato=c.sContrato and ba.sNumeroOrden=ot.sNumeroOrden) ' +
              'inner join tiposdemovimiento tm on(tm.sContrato=:Contrato and tm.sIdTipoMovimiento=ba.sIdTipoMovimiento and tm.sClasificacion="Tarifa Diaria") ' +
              'where (c.sContrato=:Contrato or c.sCodigo=:Contrato) and ba.dIdFecha=:Fecha '  +
              'group by ot.sContrato,ot.sNumeroorden';
    ParamByName('Contrato').AsString:=global_Contrato_Barco;
    ParamByName('Fecha').AsDate:=tdIdFecha.Date;
    Open;

    {while not Eof do
    begin
      dbcmbPernoctas.Items.Add(FieldByName('sDescripcionCorta').AsString);
      dbcmbPernoctas.Values.Add(FieldByName('sNumeroOrden').AsString);
      Next;
    end;

    if dbcmbPernoctas.Items.Count=0 then
    begin
      Active:=False;
      SQL.Text:='select sidpernocta,sdescripcion from pernoctan where sIdPernocta=' + QuotedStr(ordenesdetrabajo.FieldByName('sIdPernocta').AsString);
      Open;
      if RecordCount=1 then
      begin
        dbcmbPernoctas.Items.Add(FieldByName('sdescripcion').AsString);
        dbcmbPernoctas.Values.Add(FieldByName('sIdPernocta').AsString);
      end;
    end;

    if dbcmbPernoctas.Items.Count>0 then
      dbcmbPernoctas.ItemIndex:=0;   }

    while NxPgcDatos.PageCount>0 do
    begin
      sDatosPanel:=TSuperPanel(ListaObj.Objects[ListaObj.IndexOf(NxPgcDatos.Pages[0].Name)]);
      sDatosPanel.Destroy;
      NxPgcDatos.Pages[0].Destroy;
    end;

    ListaObj.Clear;
    First;
    while not eof do
    begin
      NxTbs:= TNxTabSheet.Create(Self);
      sNombre:=StringReplace(FieldByName('sNumeroOrden').AsString,' ','',[rfReplaceAll]);
      sNombre:=StringReplace(sNombre,'-','_',[rfReplaceAll]);
      NxTbs.Name:=sNombre;
      NxTbs.Caption:=FieldByName('sDescripcionCorta').AsString;
      NxPgcDatos.AddPage(NxTbs);
      sDatosPanel:=TSuperPanel.Create(NxTbs,FieldByName('sContrato').AsString,FieldByName('sNumeroOrden').AsString);
      sDatosPanel.NxGridDatos.PopupMenu:=pmPage;
      CargaDatos(FieldByName('sContrato').AsString,FieldByName('sNumeroOrden').AsString,sDatosPanel);
      //sDatosPanel.NxGridDatos.OnEnter:=MiEnter;
      ListaObj.AddObject(NxTbs.name,sDatosPanel);
      Next;
    end;

    if NxPgcDatos.PageCount>0 then
    begin
      NxPgcDatos.ActivePageIndex:=0;
      //TSuperPanel(ListaObj.Objects[ListaObj.IndexOf(NxPgcDatos.Pages[NxPgcDatos.ActivePageIndex].Name)]).NxGridDatos.SetFocus;

    end;
  end;

 { AvSPpmPernoctas.MenuItems.Clear;
  with QrPernoctas do
  begin
    Active:=False;
    SQL.Text:='select sidpernocta,sdescripcion from pernoctan ';
    Open;
    while not Eof do
    begin
      if dbcmbPernoctas.Values.IndexOf(FieldByName('sIdPernocta').AsString)=-1 then
      begin
        StMiPernocta:=AvSPpmPernoctas.MenuItems.Add;
        StMiPernocta.Caption:=FieldByName('sDescripcion').AsString;
        StMiPernocta.Checked:=False;
        StMiPernocta.Style:=isCheckBox;

      end;
      Next;
    end;
  end; }
  Recharged:=False;

end;

procedure TFrmCuadreXPartida.SdgExcelTypeChange(Sender: TObject);
begin
  case SdgExcel.FilterIndex of
    1 : SdgExcel.DefaultExt := 'xls';
    2 : SdgExcel.DefaultExt := 'xlsx';
  else
    SdgExcel.DefaultExt := '';
  end;
end;

Procedure TFrmCuadreXPartida.InicarCategoria;
begin
  ArrayTipo[1,1] := 'bitacoradepersonal'; //tabla de bitacora
  ArrayTipo[1,2] := 'personal'; //tabla de catalogo
  ArrayTipo[1,3] := 'sIdPersonal'; //Id de Categoria
  ArrayTipo[1,4] := 'Si'; //Se cobra o no la categoria,
  ArrayTipo[1,5] := 'ED'; //Tipo de movimiento en bitacora de actividades..

  ArrayTipo[2,1] := 'bitacoradeequipos'; //tabla de bitacora
  ArrayTipo[2,2] := 'equipos'; //tabla de catalogo
  ArrayTipo[2,3] := 'sIdEquipo'; //Id de Categoria
  ArrayTipo[2,4] := 'Si'; //Se cobra o no la categoria,
  ArrayTipo[2,5] := 'ED'; //Tipo de movimiento en bitacora de actividades..

  ArrayTipo[3,1] := 'bitacoradepersonal'; //tabla de bitacora
  ArrayTipo[3,2] := 'personal'; //tabla de catalogo
  ArrayTipo[3,3] := 'sIdPersonal'; //Id de Categoria
  ArrayTipo[3,4] := 'No'; //Se cobra o no la categoria,
  ArrayTipo[3,5] := 'N'; //Tipo de movimiento en bitacora de actividades..

  if tsIdCategorias.ItemIndex = 0 then
  begin
    iTipo := 1;
    sTipoFiltro:=global_labelPersonal;
  end;

  if tsIdCategorias.ItemIndex = 1 then
  begin
    iTipo := 2;
    sTipoFiltro:=global_labelEquipo;
  end;

  if tsIdCategorias.ItemIndex = 2 then
  begin
    iTipo := 3;
    sTipoFiltro:=global_labelPersonal;
  end;
end;

procedure TFrmCuadreXPartida.InsertarValoresenBlanco1Click(Sender: TObject);
var
  sDatosPanel:TSuperPanel;
  I:Integer;
  Reng:Integer;
begin
  sDatosPanel:=TSuperPanel(ListaObj.Objects[ListaObj.IndexOf(NxPgcDatos.Pages[NxPgcDatos.ActivePageIndex].Name)]);
  Reng:=0;
  with sDatosPanel.NxGridDatos do
  begin
    while Reng<RowCount do
    begin
      for I:=2 to Columns.Count-1 do
      begin
        Cell[I,Reng].AsString:='';
      end;

      sDatosPanel.NxGridTotales.Cell[0,Reng].asfloat:=0;
      Inc(reng);
    end;
  end;
end;

procedure TFrmCuadreXPartida.NxBtnEdtBuscarButtonClick(Sender: TObject);
var
  sFiltro:string;
begin
  sFiltro:='*'+NxBtnEdtBuscar.Text+'*';
  catalogo_maestro.Filtered:=False;
  catalogo_maestro.Filter:=dbUgridBuscar.Columns[0].FieldName + ' like ' + QuotedStr(sFiltro) + ' or sDescripcion like ' + QuotedStr(sFiltro);
  catalogo_maestro.Filtered:=True;
end;

procedure TFrmCuadreXPartida.NxPgcDatosChange(Sender: TObject);
begin
  if not Recharged then
    dbcmbPernoctas.ItemIndex:=NxPgcDatos.ActivePageIndex;
end;

procedure TFrmCuadreXPartida.NxPgcDatosPageClose(Sender: TObject;
  PageIndex: Integer; var Accept: Boolean);
var
  idIndex:Integer;
  sDatosPanel:TSuperPanel;
  StMiPernocta:TStickyMenuItem;
begin
  if MessageDlg('Desea eliminar el '+ArrayTipo[iTipo,2]{personal,equipo}+' del Dia '+DateToStr(tdIdFecha.Date)+
  '. Perteneciente a la Embarcacion '+ NxPgcDatos.Pages[NxPgcDatos.ActivePageIndex].Caption+ ' ?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
  begin
    idIndex:=ListaObj.IndexOf(NxPgcDatos.Pages[NxPgcDatos.ActivePageIndex].Name);
    sDatosPanel:=TSuperPanel(ListaObj.Objects[idIndex]);
    Connection.qryBusca2.SQL.Clear;
    Connection.qryBusca2.SQL.Add('Select '+ArrayTipo[iTipo,3]{sIdpersonal,sIdequipo}+' from '+ArrayTipo[iTipo,2]{personal,equipo}+' Where sContrato =:Contrato and lCobro ="'+ArrayTipo[iTipo,4]{"Si","No"}+'" ');
    Connection.qryBusca2.Params.ParamByName('Contrato').Datatype := ftString;
    Connection.qryBusca2.Params.ParamByName('Contrato').Value    := global_contrato;
    Connection.qryBusca2.Open;

    while not connection.QryBusca2.Eof do
    begin
      connection.zCommand.Active := False;
      connection.zCommand.SQL.Clear;
      connection.zCommand.SQL.Add('Delete from '+ArrayTipo[iTipo,1]{bitacoradepersonal,bitacoradeequipos}+' '+
                'where sContrato =:contrato and dIdFecha =:fecha and sidPernocta=:Pernocta and '+ArrayTipo[iTipo,3]{sIdpersonal,sIdequipo}+' =:Personal ');
      connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
      connection.zCommand.Params.ParamByName('Contrato').Value    := global_contrato;
      connection.zCommand.Params.ParamByName('Fecha').DataType    := ftDate;
      connection.zCommand.Params.ParamByName('Fecha').Value       := tdIdFecha.Date;
      connection.zCommand.Params.ParamByName('Personal').DataType := ftString;
      connection.zCommand.Params.ParamByName('Personal').Value    := Connection.qryBusca2.FieldValues[ArrayTipo[iTipo,3]{sIdpersonal,sIdequipo}];
      connection.zCommand.Params.ParamByName('Pernocta').AsString:=sDatosPanel.Id;
      connection.zCommand.ExecSQL;
      connection.QryBusca2.Next;
    end;
    sDatosPanel.NxGridDatos.ClearRows;
    sDatosPanel.ListaRecursos.Clear;
    sDatosPanel.Destroy;
    ListaObj.Delete(idIndex);

    StMiPernocta:=AvSPpmPernoctas.MenuItems.Add;
    StMiPernocta.Caption:=dbcmbPernoctas.Items.Strings[PageIndex];
    StMiPernocta.Checked:=False;
    StMiPernocta.Style:=isCheckBox;

    dbcmbPernoctas.Items.Delete(PageIndex);
    dbcmbPernoctas.Values.Delete(PageIndex);

    Accept:=True;
  end
  else
    Accept:=False;
end;

procedure TFrmCuadreXPartida.AgregarNuevaPernocta1Click(Sender: TObject);
begin
  if iTipo <> 2 then
  begin
    NGbxConsulta.Caption := '           >> CATALOGO MAESTRO DE PERSONAL';
    dbUgridBuscar.Columns[0].FieldName := 'sIdPersonal';
    dbUgridBuscar.AutoSizeRows:=true;
    btnNuevoPersonal.Enabled:=true;
  end
  else
  begin
    NGbxConsulta.Caption := '           >> CATALOGO MAESTRO DE EQUIPO';
    dbUgridBuscar.Columns[0].FieldName := 'sIdEquipo';
    dbUgridBuscar.AutoSizeRows:=False;
    dbUgridBuscar.RowResize:=True;
    dbUgridBuscar.RowsHeight:=20;
    btnNuevoPersonal.Enabled:=False;
  end;
  NGbxConsulta.Height  := 493;
  NGbxConsulta.Left    := 300;
  NGbxConsulta.Top     := 80;
  NGbxConsulta.Width   := 684;
  NGbxConsulta.Visible := True;
  catalogo_maestro.Active := False;
  catalogo_maestro.Filtered:=False;
  catalogo_maestro.SQL.Clear;
  catalogo_maestro.SQL.Add('select * from '+ArrayTipo[iTipo,2]{personal,equipo}+' where sContrato =:Contrato and lCobro = "'+ArrayTipo[iTipo,4]{"Si","No"}+'" order by iItemOrden');
  catalogo_maestro.ParamByName('Contrato').AsString := global_contrato;
  catalogo_maestro.Open;
end;

procedure TFrmCuadreXPartida.AvSPpmPernoctasButtonBar0Click(Sender: TObject);
var
  J:Integer;
  QrPernoctas:TZReadOnlyQuery;
  sDatosPanel:TSuperPanel;
  NxTbs: TNxTabSheet;
  AcPgInx:Integer;
begin
  QrPernoctas:=TZReadOnlyQuery.Create(nil);
  QrPernoctas.Connection:=connection.zConnection;
  AcPgInx:=NxPgcDatos.ActivePageIndex;
  for J := 0 to AvSPpmPernoctas.MenuItems.Count-1 do
  begin
    if AvSPpmPernoctas.MenuItems.Items[j].Checked=true then
    begin
      with QrPernoctas do
      begin
        Active:=False;
        SQL.Text:='select sidpernocta,sdescripcion from pernoctan where sDescripcion=' + QuotedStr(AvSPpmPernoctas.MenuItems.Items[j].Caption);
        Open;
        if RecordCount=1 then
        begin
          dbcmbPernoctas.Items.Add(FieldByName('sdescripcion').AsString);
          dbcmbPernoctas.Values.Add(FieldByName('sIdPernocta').AsString);

          NxTbs:= TNxTabSheet.Create(Self);
          NxTbs.Name:=StringReplace(FieldByName('sIdPernocta').AsString,' ','',[rfReplaceAll]);
          NxTbs.Caption:=FieldByName('sDescripcion').AsString;
          NxPgcDatos.AddPage(NxTbs);
          sDatosPanel:=TSuperPanel.Create(NxTbs,FieldByName('sContrato').AsString,FieldByName('sIdPernocta').AsString);
          sDatosPanel.NxGridDatos.PopupMenu:=pmPage;
          //CargaDatos(FieldByName('sIdPernocta').AsString,sDatosPanel.ListaPdas,sDatosPanel.ListaRecursos,sDatosPanel.NxGridDatos,sDatosPanel.NxGridTotales,sDatosPanel.AvGridTurnos,sDatosPanel.AvGridPdas,sDatosPanel.PSbDatos);
          //sDatosPanel.NxGridDatos.OnEnter:=MiEnter;
          ListaObj.AddObject(NxTbs.Name,sDatosPanel);

        end;
      end;
    end;
  end;

  for j := AvSPpmPernoctas.MenuItems.Count-1 downto 0 do
    if AvSPpmPernoctas.MenuItems.Items[j].Checked=true then
      AvSPpmPernoctas.MenuItems.Items[j].Destroy;

  MnIClick:=True;
  NxPgcDatos.ActivePageIndex:=NxPgcDatos.PageCount-1;
  //NxPgcDatos.Repaint;
  //TSuperPanel(ListaObj.Objects[])
end;

procedure TFrmCuadreXPartida.AvSPpmPernoctasMenuHide(Sender: TObject);
var
  J:Integer;
begin
  if not MnIClick then
  begin
    for J := 0 to AvSPpmPernoctas.MenuItems.Count-1 do
      AvSPpmPernoctas.MenuItems.Items[j].Checked:=False;
  //ShowMessage('Hide');

  end;

  MnIClick:=False;
end;

procedure TFrmCuadreXPartida.btnDeleteClick(Sender: TObject);
var
  sDatosPanel:TSuperPanel;
begin
  if MessageDlg('Desea eliminar el '+ArrayTipo[iTipo,2]{personal,equipo}+' del Dia '+DateToStr(tdIdFecha.Date)+
  '. Perteneciente a la Embarcacion '+ NxPgcDatos.Pages[NxPgcDatos.ActivePageIndex].Caption+ ' ?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
  begin
    sDatosPanel:=TSuperPanel(ListaObj.Objects[ListaObj.IndexOf(NxPgcDatos.Pages[NxPgcDatos.ActivePageIndex].Name)]);
    Connection.qryBusca2.SQL.Clear;
    Connection.qryBusca2.SQL.Add('Select '+ArrayTipo[iTipo,3]{sIdpersonal,sIdequipo}+' from '+ArrayTipo[iTipo,2]{personal,equipo}+' Where sContrato =:Contrato and lCobro ="'+ArrayTipo[iTipo,4]{"Si","No"}+'" ');
    Connection.qryBusca2.Params.ParamByName('Contrato').Datatype := ftString;
    Connection.qryBusca2.Params.ParamByName('Contrato').Value    := sDatosPanel.sContrato;
    Connection.qryBusca2.Open;

    while not connection.QryBusca2.Eof do
    begin
      connection.zCommand.Active := False;
      connection.zCommand.SQL.Clear;
      connection.zCommand.SQL.Add('Delete from '+ArrayTipo[iTipo,1]{bitacoradepersonal,bitacoradeequipos}+' '+
                'where sContrato =:contrato and dIdFecha =:fecha and '+ArrayTipo[iTipo,3]{sIdpersonal,sIdequipo}+' =:Personal ');
      connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
      connection.zCommand.Params.ParamByName('Contrato').Value    := sDatosPanel.sContrato;
      connection.zCommand.Params.ParamByName('Fecha').DataType    := ftDate;
      connection.zCommand.Params.ParamByName('Fecha').Value       := tdIdFecha.Date;
      connection.zCommand.Params.ParamByName('Personal').DataType := ftString;
      connection.zCommand.Params.ParamByName('Personal').Value    := Connection.qryBusca2.FieldValues[ArrayTipo[iTipo,3]{sIdpersonal,sIdequipo}];
     // connection.zCommand.Params.ParamByName('Pernocta').AsString:=sDatosPanel.Id;
      connection.zCommand.ExecSQL;
      connection.QryBusca2.Next;
    end;
    sDatosPanel.NxGridDatos.ClearRows;
    sDatosPanel.ListaRecursos.Clear;
  end;
end;

procedure TFrmCuadreXPartida.btnExportarClick(Sender: TObject);
const
  ExcelApp='Excel.Application';
var
  Excel:Variant;
  Libro:Variant;
  Hoja:Variant;
  sFileName:string;
  InFolder: array[0..MAX_PATH] of Char;
  pidl: PItemIDList;
  i,j:Integer;
  QrVigencia,QrMoe:TZReadOnlyQuery;
  Reng,Colum:Integer;
  QrFrentes,QrDetalle:TZReadOnlyQuery;
  TmpColum,ITotalCol,IposEnc:Integer;
  AuxCol,TmpReng:Integer;
  sPosReng,CadSuma,CadSuperSum:string;

begin
  {$REGION 'Creacion e Inicializacion de variables'}

  QrFrentes:=TZReadOnlyQuery.Create(nil);
  QrFrentes.Connection:=frm_connection.connection.zConnection;
  QrFrentes.Active:=False;
  QrFrentes.SQL.Text:='select ot.* from ordenesdetrabajo ot ' +
      'inner join contratos c on(ot.sContrato=c.sContrato) ' +
      'inner join bitacoradeactividades ba on(ba.sContrato=c.sContrato and ba.sNumeroOrden=ot.sNumeroOrden) ' +
      'inner join tiposdemovimiento tm on(tm.sContrato=:Contrato and tm.sIdTipoMovimiento=ba.sIdTipoMovimiento and tm.sClasificacion="Tarifa Diaria") ' +
      'where (c.sContrato=:Contrato or c.sCodigo=:Contrato) and ba.dIdFecha=:Fecha '  +
      'group by ot.sContrato,ot.sNumeroorden';
  QrFrentes.ParamByName('Contrato').AsString:=global_Contrato_Barco;
  QrFrentes.ParamByName('Fecha').AsDate:=tdIdFecha.Date;
  QrFrentes.Open;

  QrVigencia:=TZReadOnlyQuery.Create(nil);
  QrVigencia.Connection:=connection.zConnection;

  QrMoe:=TZReadOnlyQuery.Create(nil);
  QrMoe.Connection:=connection.zConnection;

  QrDetalle:=TZReadOnlyQuery.Create(nil);
  QrDetalle.Connection:=connection.zConnection;
  QrDetalle.SQL.Text:='select a.*,concat(a.sHoraInicio,"-",a.sHorafinal)  as Horario, tm.sIdTipoMovimiento from bitacoradeactividades a ';
  QrDetalle.SQL.add('inner join actividadesxorden o on (o.sContrato = a.sContrato and o.sIdConvenio =:convenio and o.sNumeroOrden =:orden and a.sWbs = o.sWbs and o.sNumeroActividad = a.sNumeroActividad and o.sTipoActividad = "Actividad") ');
  QrDetalle.SQL.Add('inner join tiposdemovimiento tm on(tm.sContrato=:ContratoBarco and tm.sIdTipoMovimiento=a.sIdTipoMovimiento and tm.sClasificacion="Tarifa Diaria")');
  QrDetalle.SQL.Add('where a.sContrato =:contrato and a.dIdFecha =:fecha and a.sNumeroOrden =:orden and a.sIdTurno =:turno and a.sIdTipoMovimiento = "'+ArrayTipo[iTipo,5]{"E","N"}+'" '+
                    'group by a.sContrato, a.iIdDiario '+
                    'order by a.sContrato,o.iItemOrden,a.sHoraInicio');


    // Armar el nombre de archivo
  SHGetSpecialFolderLocation(application.Handle, CSIDL_PERSONAL, pidl);
  SHGetPathFromIDList(PIDL, InFolder);
  sFileName := InFolder;

  if sFileName[Length(sFileName)] <> '\' then
    sFileName := sFileName + '\';

  SdgExcel.InitialDir:= sFileName;
  SdgExcel.FileName:='Cuadre Personal-Equipo ' + IntToStr(YearOf(tdIdFecha.Date)) + '-' + IntToStr(MonthOf(tdIdFecha.Date)) + '-' + IntToStr(DayOf(tdIdFecha.Date)) + '.xlsx';



  {$ENDREGION}

  {$REGION 'Proceso de Exportacion'}
  if SdgExcel.Execute then
  begin
    try
      Excel := CreateOleObject('Excel.Application');
    except
      FreeAndNil(Excel);
      showmessage('No es posible generar el ambiente de EXCEL, informe de esto al administrador del sistema.');
      Exit;
    end;

    try
      Excel.Visible := False;
      Excel.DisplayAlerts := False;
      Excel.ScreenUpdating := False;

      Libro := Excel.Workbooks.Add;

      while Libro.Sheets.Count < 2 do
        Libro.Sheets.Add;

      while Libro.Sheets.Count > 2 do
        Excel.ActiveWindow.SelectedSheets.Delete;

      for I := 1 to 2 do
      begin
        Hoja:=Libro.Sheets[i];
        Hoja.Select;
        
        if I=1 then
          Hoja.Name:='MANO DE OBRA'
        else
          Hoja.Name:='EQUIPO';

        Excel.ActiveWindow.Zoom := 70;
        sPosReng:='';
        CadSuperSum:='';
        Reng:=4;
        QrMoe.Active:=False;
        QrMoe.SQL.Text:='select mr.* from moerecursos mr inner join moe m on(m.iidmoe=mr.iidmoe) ' +
                        'where m.sContrato=:Contrato and m.didfecha=(select max(dIdFecha) from moe '+
                        'where scontrato=m.scontrato and didfecha<=:Fecha) and mr.eTipoRecurso=:Tipo';
        QrMoe.ParamByName('Contrato').AsString:=global_Contrato_Barco;
        QrMoe.ParamByName('Fecha').AsDate:=tdIdFecha.Date;
        if i=1 then
          QrMoe.ParamByName('Tipo').AsString:='Personal'
        else
          QrMoe.ParamByName('Tipo').AsString:='Equipo';
        QrMoe.Open;

        if QrMoe.RecordCount>0 then
        begin
          Colum:=1;
          Excel.Columns[ColumnaNombre(Colum)+':'+ColumnaNombre(Colum)].ColumnWidth :=20;

          Hoja.Range[ColumnaNombre(Colum)+'1:'+columnanombre(Colum) + '1'].Select;
          Excel.Selection.Value:=I;

          inc(Colum);
          Excel.Columns[ColumnaNombre(Colum)+':'+ColumnaNombre(Colum)].ColumnWidth :=40;
          Excel.Rows[IntToStr(Reng) + ':' + IntToStr(Reng)].RowHeight :=45;
          Hoja.Range[ColumnaNombre(Colum)+IntToStr(Reng) + ':'+ColumnaNombre(Colum)+IntToStr(Reng)].select;
          Excel.Selection.NumberFormat := '[$-F800]dddd, mmmm dd, yyyy' ;
          Excel.Selection.Value:=TdidFecha.Date;

          inc(Colum);
          Excel.Columns[ColumnaNombre(Colum)+':'+ColumnaNombre(Colum)].select;
          Excel.Selection.EntireColumn.Hidden := True;

          inc(Colum,1);
          Excel.Columns[ColumnaNombre(Colum)+':'+ColumnaNombre(Colum)].ColumnWidth :=10;

          inc(Colum,1);
          Excel.Columns[ColumnaNombre(Colum)+':'+ColumnaNombre(Colum)].ColumnWidth :=7;
          Hoja.Range[ColumnaNombre(Colum)+IntToStr(Reng) + ':'+ColumnaNombre(Colum+1)+IntToStr(Reng+1)].select;
          Excel.Selection.MergeCells := True;
          Excel.Selection.Font.Name := 'Arial';
          Excel.Selection.NumberFormat := '@' ;
          Excel.Selection.Value :='TOTAL';
          Excel.Selection.Font.Bold := True;
          Excel.Selection.HorizontalAlignment := xlCenter;
          Excel.Selection.VerticalAlignment := xlCenter;
          Excel.Selection.Font.Color:=xlBlueKing;

          inc(Colum);
          Excel.Columns[ColumnaNombre(Colum)+':'+ColumnaNombre(Colum)].ColumnWidth :=12;

          inc(Colum,1);
          Excel.Columns[ColumnaNombre(Colum)+':'+ColumnaNombre(Colum)].ColumnWidth :=12;
          Hoja.Range[ColumnaNombre(Colum)+IntToStr(Reng) + ':'+ColumnaNombre(Colum+1)+IntToStr(Reng+1)].select;
          Excel.Selection.MergeCells := True;
          Excel.Selection.Font.Name := 'Arial';
          Excel.Selection.NumberFormat := '#0.00' ;
          Excel.Selection.Value :=0;
          Excel.Selection.Font.Color:=xlBlueKing;
          Excel.Selection.Font.Bold := True;
          Excel.Selection.HorizontalAlignment := xlCenter;
          Excel.Selection.VerticalAlignment := xlCenter;
          Excel.Selection.Interior.Color:=255;

          inc(Colum,1);
          Excel.Columns[ColumnaNombre(Colum)+':'+ColumnaNombre(Colum)].ColumnWidth :=15;

          Inc(Colum,1);
          Excel.Columns[ColumnaNombre(Colum)+':'+ColumnaNombre(Colum)].ColumnWidth :=8;
          TmpColum:=Colum;
          Qrmoe.First;
          while Not QrMoe.Eof do
          begin
            Hoja.Range[ColumnaNombre(Colum)+IntToStr(Reng) + ':'+ColumnaNombre(Colum+1)+IntToStr(Reng)].select;
            Excel.Selection.MergeCells := True;
            Excel.Selection.Font.Name := 'Arial';
            Excel.Selection.Font.size:=9;
            Excel.Selection.HorizontalAlignment := xlRight;
            Excel.Selection.VerticalAlignment := xlTop;
            Excel.Selection.NumberFormat := '@' ;
            Excel.Selection.Value :=QrMoe.FieldByName('sDescripcion').AsString;
            Excel.Selection.wraptext:=True;

            Hoja.Range[ColumnaNombre(Colum)+IntToStr(Reng+1) + ':'+ColumnaNombre(Colum)+IntToStr(Reng+1)].select;
            Excel.Selection.MergeCells := True;
            Excel.Selection.Font.Name := 'Arial';
            Excel.Selection.NumberFormat := '@' ;
            Excel.Selection.Font.size:=11;
            Excel.Selection.HorizontalAlignment := xlCenter;
            Excel.Selection.VerticalAlignment := xlCenter;
            Excel.Selection.Font.Color:=xlBlueKing;
            Excel.Selection.Value :=QrMoe.FieldByName('sIdRecurso').AsString;
            Excel.Selection.wraptext:=True;

            Hoja.Range[ColumnaNombre(Colum)+IntToStr(Reng+2) + ':'+ColumnaNombre(Colum)+IntToStr(Reng+2)].select;
            Excel.Selection.MergeCells := True;
            Excel.Selection.Font.Name := 'Arial';
            Excel.Selection.NumberFormat := '#0.00' ;
            Excel.Selection.Value :=QrMoe.FieldByName('dCantidad').AsFloat;
            Excel.Selection.wraptext:=True;

            Hoja.Range[ColumnaNombre(Colum+1)+IntToStr(Reng+2) + ':'+ColumnaNombre(Colum+1)+IntToStr(Reng+2)].select;
            Excel.Selection.MergeCells := True;
            Excel.Selection.Font.Name := 'Arial';
            //Excel.Selection.NumberFormat := '#0.00' ;
            Excel.Selection.Value :=0;
            Excel.Selection.wraptext:=True;

            Inc(Colum,1);
            Excel.Columns[ColumnaNombre(Colum)+':'+ColumnaNombre(Colum)].ColumnWidth :=13;

            Inc(Colum,1);
            Excel.Columns[ColumnaNombre(Colum)+':'+ColumnaNombre(Colum)].ColumnWidth :=8;
            QrMoe.Next;
          end;
          IposEnc:=Reng+2;
          Inc(Reng,3);
          QrFrentes.First;
          while not QrFrentes.Eof do
          begin
            Excel.Rows[IntToStr(Reng) + ':' + IntToStr(Reng)].RowHeight :=45;

            Colum:= TmpColum;
            Qrmoe.First;
            while Not QrMoe.Eof do
            begin
              Hoja.Range[ColumnaNombre(Colum)+IntToStr(Reng) + ':'+ColumnaNombre(Colum+1)+IntToStr(Reng)].select;
              Excel.Selection.MergeCells := True;
              Excel.Selection.Font.Name := 'Arial';
              Excel.Selection.NumberFormat := '@' ;
              Excel.Selection.Font.size:=9;
              Excel.Selection.HorizontalAlignment := xlRight;
              Excel.Selection.VerticalAlignment := xlTop;
              Excel.Selection.Value :=QrMoe.FieldByName('sDescripcion').AsString;
              Excel.Selection.wraptext:=True;

              Hoja.Range[ColumnaNombre(Colum)+IntToStr(Reng+1) + ':'+ColumnaNombre(Colum)+IntToStr(Reng+1)].select;
              Excel.Selection.MergeCells := True;
              Excel.Selection.Font.Name := 'Arial';
              Excel.Selection.NumberFormat := '@' ;
              Excel.Selection.Value :='Cant.';
              Excel.Selection.HorizontalAlignment := xlCenter;
              Excel.Selection.VerticalAlignment := xlCenter;
              Excel.Selection.wraptext:=True;

              Hoja.Range[ColumnaNombre(Colum+1)+IntToStr(Reng+1) + ':'+ColumnaNombre(Colum+1)+IntToStr(Reng+1)].select;
              Excel.Selection.MergeCells := True;
              Excel.Selection.Font.Name := 'Arial';
              Excel.Selection.NumberFormat := '@' ;
              Excel.Selection.Value :='H.H';
              Excel.Selection.HorizontalAlignment := xlCenter;
              Excel.Selection.VerticalAlignment := xlCenter;
              Excel.Selection.wraptext:=True;
              
              Inc(Colum,2);
              QrMoe.Next;
            end;

            Inc(reng);
            Hoja.Range[ColumnaNombre(2)+IntToStr(Reng) + ':'+ColumnaNombre(TmpColum-1)+IntToStr(Reng)].select;
            Excel.Selection.MergeCells := True;
            Excel.Selection.Font.Name := 'Arial';
            Excel.Selection.NumberFormat := '@' ;
            Excel.Selection.Font.size:=10;
            Excel.Selection.Font.Bold := True;
            Excel.Selection.HorizontalAlignment := xlCenter;
            Excel.Selection.VerticalAlignment := xlCenter;
            Excel.Selection.Value :=QrFrentes.FieldByName('sDescripcionCorta').AsString;
            Excel.Selection.wraptext:=True;
            Excel.Selection.Font.Color:=xlBlueKing;

            QrDetalle.Active:=False;
            QrDetalle.ParamByName('Contrato').AsString := QrFrentes.FieldByName('sContrato').AsString;
            QrDetalle.ParamByName('ContratoBarco').AsString := global_Contrato_Barco;
            QrDetalle.ParamByName('Convenio').AsString := global_convenio;
            QrDetalle.ParamByName('Orden').AsString    := QrFrentes.FieldByName('sNumeroOrden').AsString;
            QrDetalle.ParamByName('Turno').AsString    := global_turno;
            QrDetalle.ParamByName('fecha').AsDate      := tdIdFecha.Date;
            QrDetalle.Open;

            ITotalCol:=(TmpColum-1) + (QrMoe.RecordCount*2);

            Inc(Reng,2);
            TmpReng:=Reng;
            while not QrDetalle.Eof do
            begin
              Colum:=TmpColum-8;
              Hoja.Range[ColumnaNombre(Colum)+IntToStr(Reng) + ':'+ColumnaNombre(Colum)+IntToStr(Reng)].select;
              Excel.Selection.MergeCells := True;
              Excel.Selection.Font.Name := 'Arial';
              Excel.Selection.NumberFormat := '@' ;
              Excel.Selection.Value :=QrFrentes.FieldByName('sNumeroOrden').AsString;
              Excel.Selection.wraptext:=True;

              Inc(Colum);
              Hoja.Range[ColumnaNombre(Colum)+IntToStr(Reng) + ':'+ColumnaNombre(Colum)+IntToStr(Reng)].select;
              Excel.Selection.MergeCells := True;
              Excel.Selection.Font.Name := 'Arial';
              Excel.Selection.NumberFormat := '@' ;
              Excel.Selection.Font.size:=10;
              Excel.Selection.Value :=QrDetalle.FieldByName('mDescripcion').AsString;
              Excel.Selection.wraptext:=True;      

              Inc(Colum);
              Hoja.Range[ColumnaNombre(Colum)+IntToStr(Reng) + ':'+ColumnaNombre(Colum)+IntToStr(Reng)].select;
              Excel.Selection.MergeCells := True;
              Excel.Selection.Font.Name := 'Arial';
              Excel.Selection.NumberFormat := '@' ;
              Excel.Selection.Value :=QrDetalle.FieldByName('iIdDiario').AsString;
              Excel.Selection.wraptext:=True;

              Inc(Colum);
              Hoja.Range[ColumnaNombre(Colum)+IntToStr(Reng) + ':'+ColumnaNombre(Colum)+IntToStr(Reng)].select;
              Excel.Selection.MergeCells := True;
              Excel.Selection.Font.Name := 'Arial';
              Excel.Selection.NumberFormat := '@' ;
              Excel.Selection.Value :=QrDetalle.FieldByName('sNumeroActividad').AsString;
              Excel.Selection.wraptext:=True;

              Inc(Colum);
              Hoja.Range[ColumnaNombre(Colum)+IntToStr(Reng) + ':'+ColumnaNombre(Colum)+IntToStr(Reng)].select;
              Excel.Selection.MergeCells := True;
              Excel.Selection.Font.Name := 'Arial';
              Excel.Selection.NumberFormat := '@' ;
              Excel.Selection.Value :=QrDetalle.FieldByName('sIdClasificacion').AsString;
              Excel.Selection.wraptext:=True;

              inc(Colum);
              Hoja.Range[ColumnaNombre(Colum)+IntToStr(Reng) + ':'+ColumnaNombre(Colum)+IntToStr(Reng)].select;
              Excel.Selection.MergeCells := True;
              Excel.Selection.Font.Name := 'Arial';
              Excel.Selection.NumberFormat := '@' ;
              Excel.Selection.Value :=QrDetalle.FieldByName('sHoraInicio').AsString;
              Excel.Selection.wraptext:=True;

              Inc(Colum);
              Hoja.Range[ColumnaNombre(Colum)+IntToStr(Reng) + ':'+ColumnaNombre(Colum)+IntToStr(Reng)].select;
              Excel.Selection.MergeCells := True;
              Excel.Selection.Font.Name := 'Arial';
              Excel.Selection.NumberFormat := '@' ;
              Excel.Selection.Value :=QrDetalle.FieldByName('sHoraFinal').AsString;
              Excel.Selection.wraptext:=True;

              Inc(Colum);
              Hoja.Range[ColumnaNombre(Colum)+IntToStr(Reng) + ':'+ColumnaNombre(Colum)+IntToStr(Reng)].select;
              Excel.Selection.Font.Name := 'Arial';
              Excel.Selection.FormulaR1C1 := '=RC[-1]-RC[-2]';
              AuxCol := TmpColum-1;
              CadSuma:='';
              while AuxCol<ITotalCol do
              begin
                Inc(AuxCol,1);
                Hoja.Range[ColumnaNombre(AuxCol)+IntToStr(Reng)].select;
                Excel.Selection.Locked := False;
                Excel.Selection.FormulaHidden := False;

                Inc(AuxCol,1);
                Hoja.Range[ColumnaNombre(AuxCol)+IntToStr(Reng)].select;
                if I=1 then
                  Excel.Selection.FormulaR1C1 :='=(RC[-1]*RC['+'-'+ inttostr(AuxCol-(TmpColum-1))+'])*2'
                else
                  Excel.Selection.FormulaR1C1 :='=(RC[-1]*RC['+'-'+ inttostr(AuxCol-(TmpColum-1))+'])';
                  
                if CadSuma='' then
                  CadSuma:='=RC[-' + IntToStr((ITotalCol+1)-AuxCol) +']'
                else
                  CadSuma:=CadSuma + '+RC[-' + IntToStr((ITotalCol+1)-AuxCol) +']';
              
              end;
              Hoja.Range[ColumnaNombre(ITotalCol+1)+IntToStr(Reng)].select;
              Excel.Selection.FormulaR1C1 :=CadSuma;

              inc(Reng);
              QrDetalle.Next;
            end;

            AuxCol := TmpColum-1;
            while AuxCol<ITotalCol do
            begin
              Inc(AuxCol,2);
              Hoja.Range[ColumnaNombre(AuxCol)+IntToStr(Reng)].select;
              Excel.Selection.Font.Bold := True;
              Excel.Selection.FormulaR1C1 :='=SUM(R[-'+Inttostr(reng-Tmpreng)+']C:R[-1]C)'
            end;

            Hoja.Range[ColumnaNombre(ITotalCol+1)+IntToStr(Reng)].select;
            Excel.Selection.Font.Bold := True;
            Excel.Selection.FormulaR1C1 :='=SUM(R[-'+Inttostr(reng-Tmpreng)+']C:R[-1]C)';

            if CadSuperSum='' then
              CadSuperSum:='=R[' + inttostr(reng-(IposEnc-2)) + ']C[' + IntToStr((ITotalCol+1)-(TmpColum-2))+']'
            else
              CadSuperSum:=CadSuperSum + '+R[' + inttostr(reng-(IposEnc-2)) + ']C[' + IntToStr((ITotalCol+1)-(TmpColum-2))+']';

            if sPosReng='' then
              sPosReng:=inttostr(Reng)
            else
              sPosReng:=sPosReng+'-'+inttostr(Reng);

            inc(Reng,4);
            QrFrentes.Next;
          end;

          Dec(Reng,2);
          AuxCol := TmpColum-1;
          while AuxCol<ITotalCol do
          begin
            Inc(AuxCol,2);
            Hoja.Range[ColumnaNombre(AuxCol)+IntToStr(Reng)].select;
            Excel.Selection.Font.Bold := True;
            Excel.Selection.NumberFormat :='0.########';
            CadSuma:='';
            for j := 1 to NumItems(sPosReng,'-') do
            begin
              if CadSuma='' then
                CadSuma:='=R[-' + IntToStr(reng-strtoint(TraerItem(sPosReng,'-',j)))+']C'
              else
                CadSuma:=CadSuma + '+' + 'R[-' + IntToStr(reng-strtoint(TraerItem(sPosReng,'-',j)))+']C';
            end;
            Excel.Selection.FormulaR1C1 :=CadSuma;

            Hoja.Range[ColumnaNombre(AuxCol)+IntToStr(IposEnc)].select;
            Excel.Selection.Font.Bold := True;
            Excel.Selection.FormulaR1C1 :='=R[' + IntToStr(reng-iPosenc)+ ']C';
          end;

          Hoja.Range[ColumnaNombre(ITotalCol+1)+IntToStr(Reng)].select;
          Excel.Selection.FormulaR1C1 :=CadSuma;

          Hoja.Range[ColumnaNombre(TmpColum-2)+IntToStr(IposEnc-2) + ':' + ColumnaNombre(TmpColum-1)+IntToStr(IposEnc-1)].select;
          Excel.Selection.Font.Bold := True;
          Excel.Selection.FormulaR1C1 :=CadSuperSum;

          AuxCol := TmpColum-1;
          while AuxCol<ITotalCol do
          begin
            Inc(AuxCol,2);
            Excel.ActiveSheet.Range[ColumnaNombre(AuxCol-1)+IntToStr(IposEnc)].select;
            Excel.Selection.FormatConditions.Delete;
            Excel.Selection.FormatConditions.Add(xlCellValue, xlLess, '=' + ColumnaNombre(AuxCol) + IntToStr(Reng));
            Excel.Selection.FormatConditions[1].Font.ColorIndex := 6;
            Excel.Selection.FormatConditions[1].Interior.ColorIndex := 3;
            Excel.Selection.FormatConditions.Add(xlCellValue, xlEqual, '=' +  ColumnaNombre(AuxCol) + IntToStr(Reng));
            Excel.Selection.FormatConditions[2].Font.ColorIndex := 2;
            Excel.Selection.FormatConditions[2].Interior.ColorIndex := 5;
            Excel.Selection.FormatConditions.Add(xlCellValue, xlGreater, '=' + ColumnaNombre(AuxCol) + IntToStr(Reng));
            Excel.Selection.FormatConditions[3].Interior.ColorIndex := 10;
          end;

         (* QrVigencia.Active:=False;
          QrVigencia.SQL.Text:= 'select mro.*,p.*,ot.sIdPlataforma from movtorecursosxoficio mro inner join ordenesdetrabajogral otg ' +
                                'on (mro.sContrato=otg.sContrato and mro.sNumeroOrden=otg.sNumeroOrden and mro.dFechaVigencia=otg.dFechaVigencia and mro.IFolioOficio=otg.IFolioOficio) ' +
                                'inner join '+ArrayTipo[i,2]{personal,equipo}+' p on (p.sContrato = otg.sContrato and p.'+ArrayTipo[i,3]{sIdpersonal,sIdequipo}+' = mro.sNumeroActividad'{sIdpersonal,sIdequipo}+' and p.lCobro ="'+ArrayTipo[i,4]{"Si","No"}+'") '+
                                'inner join ordenesdetrabajo ot on(ot.scontrato=otg.sContrato and ot.sNumeroOrden=otg.snumeroorden)'+
                                'where otg.dFiProgramado=' +
                                '(select max(otg2.dFiProgramado) from ordenesdetrabajogral otg2 where otg2.sContrato=otg.sContrato and otg2.sNumeroOrden=otg.sNumeroOrden ' +
                                'and otg2.dFiProgramado <=:Fecha) and mro.sAnexo=:Tipo' +
                                'order by p.iitemorden group by p.'+ArrayTipo[i,3]; *)



        end;
//        Excel.ActiveSheet.Protect('$SISADM', False);

      end;
    finally
      Excel.ActiveWorkbook.SaveAs(SdgExcel.FileName);
      Excel.Visible := True;
      Excel.DisplayAlerts := True;
      Excel.ScreenUpdating := True;
    end;

  end;
  {$ENDREGION}
end;

function TFrmCuadreXPartida.ValidarCategoria(var Excel:Variant; var Hoja:Variant;sTipo:string;dParamFecha:TDate):Boolean;
var
  ren,RenId:Integer;
  sDato:string;
  ColInicio,iSaltos,iTipoHoja:Integer;
  Encontrado,salir:Boolean;
  sPartida,sDCant,sTFiltro:string;
  Colum:Integer;
  sFolio,sIdRecurso,sParamContrato:string;
  dCantSol,dCantMoe:Extended;
  QrBusca,QrVigencia:TZReadOnlyQuery;
  Error:Boolean;
  bErr1,bErr2,bErr3:Boolean;
begin

  Error:=False;
  QrBusca:=TZReadOnlyQuery.Create(nil);
  QrBusca.Connection:=connection.zConnection;

  QrVigencia:=TZReadOnlyQuery.Create(nil);
  QrVigencia.Connection:=connection.zConnection;

  ColInicio:=9;
  ren:=8;
  RenId:=5;
  iSaltos:=0;

  bErr1:=False;
  bErr2:=False;
  bErr3:=False;

  if sTipo='Personal' then
    iTipoHoja:=1;

  if sTipo='Equipo' then
    iTipoHoja:=2;

   if iTipoHoja=1 then
      sTFiltro:=global_labelPersonal;

    if iTipoHoja=2 then
      sTFiltro:=global_labelEquipo;




  Colum:=ColInicio;
  repeat
    sDCant:=Hoja.Cells[RenId+1,Colum+1].value;
    dCantSol:=StrToFloatDef(sDCant,0);

    sDCant:=Hoja.Cells[RenId+1,Colum].value;
    dCantMoe:=StrToFloatDef(sDCant,0);

    if dCantSol<>dCantMoe then
    begin
      Hoja.Range[ColumnaNombre(Colum+1)+IntToStr(RenId+1) + ':'+ColumnaNombre(Colum+1)+IntToStr(RenId+1)].select;
      Excel.Selection.Interior.ColorIndex := 3;
      Excel.Selection.Interior.Pattern := xlSolid;
      Error:=True;
      bErr1:=True;
    end;
    Inc(Colum,2);
  until (Hoja.Cells[RenId,Colum].value='');


  sDato:=Hoja.Cells[Ren, 1].value;
  salir:=False;

  while NOt salir do
  begin
    Encontrado:=False;
    while not Encontrado do
    begin
      sDato:=Hoja.Cells[Ren, 1].value;
      if Trim(sDato)<>'' then
        Encontrado:=True
      else
      begin
        Inc(iSaltos);
        Inc(ren);
      end;

      if Encontrado then
      begin
        QrBusca.Active:=False;
        QrBusca.SQL.Text:='select c.scontrato from contratos c inner join ordenesdetrabajo ot '+
                          'on (ot.scontrato=c.scontrato) where c.scodigo=:ContratoBarco and ' +
                          'ot.snumeroorden=:Orden' ;
        QrBusca.ParamByName('ContratoBarco').AsString:=global_Contrato_Barco;
        QrBusca.ParamByName('Orden').AsString:=Trim(sDato);
        QrBusca.Open;

        if QrBusca.RecordCount=1 then
          sParamContrato:=QrBusca.FieldByName('sContrato').AsString;
      end;

      if iSaltos=8 then
      begin
        Salir:=True;
        Break;//Encontrado:=True;
      end;
    end;

    if Encontrado then
    begin
      repeat
        sFolio:=Hoja.Cells[Ren, 1].value;;
        sPartida:=Hoja.Cells[Ren, 4].value;
        Colum:=ColInicio;
        repeat
          sIdRecurso:=Hoja.Cells[RenId,Colum].value;
          sDCant:=Hoja.Cells[Ren,Colum].value;
          dCantSol:=StrToFloatDef(sDCant,0);
          if dCantSol<>0 then
          begin
            QrVigencia.Active:=False;
            QrVigencia.SQL.Text:= 'select mro.*,p.*,ot.sIdPlataforma from movtorecursosxoficio mro inner join ordenesdetrabajogral otg ' +
                                  'on (mro.sContrato=otg.sContrato and mro.sNumeroOrden=otg.sNumeroOrden and mro.dFechaVigencia=otg.dFechaVigencia and mro.IFolioOficio=otg.IFolioOficio) ' +
                                  'inner join '+ArrayTipo[iTipoHoja,2]{personal,equipo}+' p on (p.sContrato = otg.sContrato and p.'+ArrayTipo[iTipoHoja,3]{sIdpersonal,sIdequipo}+' = mro.sNumeroActividad'{sIdpersonal,sIdequipo}+' and p.lCobro ="'+ArrayTipo[iTipoHoja,4]{"Si","No"}+'") '+
                                  'inner join ordenesdetrabajo ot on(ot.scontrato=otg.sContrato and ot.sNumeroOrden=otg.snumeroorden)'+
                                  'where otg.sContrato=:Contrato and otg.sNumeroOrden=:Orden and otg.dFiProgramado=' +
                                  '(select max(otg2.dFiProgramado) from ordenesdetrabajogral otg2 where otg2.sContrato=otg.sContrato and otg2.sNumeroOrden=otg.sNumeroOrden ' +
                                  'and otg2.dFiProgramado <=:Fecha) and mro.sAnexo=:Tipo and mro.snumeroactividad=:Id';
            QrVigencia.ParamByName('Contrato').AsString:=sParamContrato;
            QrVigencia.ParamByName('Orden').AsString:=sFolio;
            QrVigencia.ParamByName('Fecha').AsDate:=dParamFecha;
            QrVigencia.ParamByName('Tipo').AsString:=sTFiltro;
            QrVigencia.ParamByName('Id').AsString:=sIdRecurso;
            QrVigencia.Open;

            if QrVigencia.recordcount=0 then
            begin
             // Hoja.Range[ColumnaNombre(Colum) + IntToStr(ren)].Select;
             Hoja.Range[ColumnaNombre(Colum)+IntToStr(Ren) + ':'+ColumnaNombre(Colum)+IntToStr(Ren)].select;
              Excel.Selection.Interior.ColorIndex := 3;
              Excel.Selection.Interior.Pattern := xlSolid;
              //Excel.Selection.Interior.Color:=3;
              Error:=True;
              bErr2:=True;
            end;

            sDCant:=Hoja.Cells[RenId+1,Colum].value;
            dCantMoe:=StrToFloatDef(sDCant,0);

            if dCantMoe<dCantSol then
            begin
              Hoja.Range[ColumnaNombre(Colum)+IntToStr(Ren) + ':'+ColumnaNombre(Colum)+IntToStr(Ren)].select;

              Excel.Selection.Borders[xlEdgeLeft].LineStyle := xlContinuous;
              Excel.Selection.Borders[xlEdgeLeft].Color := -16711681;
              Excel.Selection.Borders[xlEdgeLeft].TintAndShade := 0;
              Excel.Selection.Borders[xlEdgeLeft].Weight := xlMedium;

              Excel.Selection.Borders[xlEdgeTop].LineStyle := xlContinuous;
              Excel.Selection.Borders[xlEdgeTop].Color := -16711681;
              Excel.Selection.Borders[xlEdgeTop].TintAndShade := 0;
              Excel.Selection.Borders[xlEdgeTop].Weight := xlMedium;

              Excel.Selection.Borders[xlEdgeBottom].LineStyle := xlContinuous;
              Excel.Selection.Borders[xlEdgeBottom].Color := -16711681;
              Excel.Selection.Borders[xlEdgeBottom].TintAndShade := 0;
              Excel.Selection.Borders[xlEdgeBottom].Weight := xlMedium;

              Excel.Selection.Borders[xlEdgeRight].LineStyle := xlContinuous;
              Excel.Selection.Borders[xlEdgeRight].Color := -16711681;
              Excel.Selection.Borders[xlEdgeRight].TintAndShade := 0;
              Excel.Selection.Borders[xlEdgeRight].Weight := xlMedium;
              Error:=True;
              bErr3:=True;
            end;
          end;
          Inc(Colum,2);
        until (Hoja.Cells[RenId,Colum].value='');
        Inc(ren);
      until (Hoja.Cells[Ren, 1].value='');
    end;




  end;

  if (bErr1) or (bErr2) or (bErr3) then
  begin

    ren:=RenId;
    inc(Colum,3);
    Excel.Columns[ColumnaNombre(Colum)+ ':'+ColumnaNombre(Colum)].ColumnWidth :=7;
    if (bErr1) or (bErr2) then
    begin
      Hoja.Range[ColumnaNombre(Colum)+IntToStr(Ren) + ':'+ColumnaNombre(Colum)+IntToStr(Ren+1)].select;
      Excel.Selection.MergeCells := True;
      Excel.Selection.Interior.ColorIndex := 3;
      Excel.Selection.Interior.Pattern := xlSolid;
      Excel.Selection.MergeCells := True;
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.VerticalAlignment := xlCenter;
      Excel.Selection.Font.Size :=12;
      Excel.Selection.Font.Bold := True;
      Excel.Selection.Font.Name := 'Arial';
      Excel.Selection.wraptext:=True;

      Hoja.Range[ColumnaNombre(Colum+1)+IntToStr(Ren) + ':'+ColumnaNombre(Colum+10)+IntToStr(Ren+1)].select;
      Excel.Selection.MergeCells := True;
      Excel.Selection.HorizontalAlignment := xlLeft;
      Excel.Selection.VerticalAlignment := xlCenter;
      Excel.Selection.Font.Size :=12;
      Excel.Selection.Font.Bold := True;
      Excel.Selection.Font.Name := 'Arial';
      Excel.Selection.wraptext:=True;
      Excel.Selection.Value :='Cantidad Capturada es Diferente a la Cantidad Solicitada o Categoria no encontrada en la vigencia de ese Folio.';

      Inc(Ren,2);
    end;

    if (bErr3) then
    begin
      if ren=RenId then
      begin
        Hoja.Range[ColumnaNombre(Colum)+IntToStr(Ren) + ':'+ColumnaNombre(Colum)+IntToStr(Ren+1)].select;
        Excel.Selection.MergeCells := True;

        Excel.Selection.Borders[xlEdgeLeft].LineStyle := xlContinuous;
        Excel.Selection.Borders[xlEdgeLeft].Color := -16711681;
        Excel.Selection.Borders[xlEdgeLeft].TintAndShade := 0;
        Excel.Selection.Borders[xlEdgeLeft].Weight := xlMedium;

        Excel.Selection.Borders[xlEdgeTop].LineStyle := xlContinuous;
        Excel.Selection.Borders[xlEdgeTop].Color := -16711681;
        Excel.Selection.Borders[xlEdgeTop].TintAndShade := 0;
        Excel.Selection.Borders[xlEdgeTop].Weight := xlMedium;

        Excel.Selection.Borders[xlEdgeBottom].LineStyle := xlContinuous;
        Excel.Selection.Borders[xlEdgeBottom].Color := -16711681;
        Excel.Selection.Borders[xlEdgeBottom].TintAndShade := 0;
        Excel.Selection.Borders[xlEdgeBottom].Weight := xlMedium;

        Excel.Selection.Borders[xlEdgeRight].LineStyle := xlContinuous;
        Excel.Selection.Borders[xlEdgeRight].Color := -16711681;
        Excel.Selection.Borders[xlEdgeRight].TintAndShade := 0;
        Excel.Selection.Borders[xlEdgeRight].Weight := xlMedium;

        Hoja.Range[ColumnaNombre(Colum+1)+IntToStr(Ren) + ':'+ColumnaNombre(Colum+10)+IntToStr(Ren+1)].select;
        Excel.Selection.MergeCells := True;
        Excel.Selection.HorizontalAlignment := xlLeft;
        Excel.Selection.VerticalAlignment := xlCenter;
        Excel.Selection.Font.Size :=12;
        Excel.Selection.Font.Bold := True;
        Excel.Selection.Font.Name := 'Arial';
        Excel.Selection.wraptext:=True;
        Excel.Selection.Value :='Cantidad Capturada es Mayor a la Cantidad A Bordo de la Embarcacion.';
      end
      else
      begin
        Hoja.Range[ColumnaNombre(Colum)+IntToStr(Ren) + ':'+ColumnaNombre(Colum)+IntToStr(Ren)].select;
        Excel.Selection.MergeCells := True;

        Excel.Selection.Borders[xlEdgeLeft].LineStyle := xlContinuous;
        Excel.Selection.Borders[xlEdgeLeft].Color := -16711681;
        Excel.Selection.Borders[xlEdgeLeft].TintAndShade := 0;
        Excel.Selection.Borders[xlEdgeLeft].Weight := xlMedium;

        Excel.Selection.Borders[xlEdgeTop].LineStyle := xlContinuous;
        Excel.Selection.Borders[xlEdgeTop].Color := -16711681;
        Excel.Selection.Borders[xlEdgeTop].TintAndShade := 0;
        Excel.Selection.Borders[xlEdgeTop].Weight := xlMedium;

        Excel.Selection.Borders[xlEdgeBottom].LineStyle := xlContinuous;
        Excel.Selection.Borders[xlEdgeBottom].Color := -16711681;
        Excel.Selection.Borders[xlEdgeBottom].TintAndShade := 0;
        Excel.Selection.Borders[xlEdgeBottom].Weight := xlMedium;

        Excel.Selection.Borders[xlEdgeRight].LineStyle := xlContinuous;
        Excel.Selection.Borders[xlEdgeRight].Color := -16711681;
        Excel.Selection.Borders[xlEdgeRight].TintAndShade := 0;
        Excel.Selection.Borders[xlEdgeRight].Weight := xlMedium;

        Hoja.Range[ColumnaNombre(Colum+1)+IntToStr(Ren) + ':'+ColumnaNombre(Colum+10)+IntToStr(Ren)].select;
        Excel.Selection.MergeCells := True;
        Excel.Selection.HorizontalAlignment := xlLeft;
        Excel.Selection.VerticalAlignment := xlCenter;
        Excel.Selection.Font.Size :=12;
        Excel.Selection.Font.Bold := True;
        Excel.Selection.Font.Name := 'Arial';
        Excel.Selection.wraptext:=True;
        Excel.Selection.Value :='Cantidad Capturada es Mayor a la Cantidad A Bordo de la Embarcacion.';
      end;

      Inc(Ren);
    end;


  end;


  Result:=Error;
////Aqui va todo
end;

procedure TFrmCuadreXPartida.btnImportarClick(Sender: TObject);
const
  ExcelApp='Excel.Application';
var
  Excel:Variant;
  Libro:Variant;
  Hoja:Variant;
  I:Integer;
  RengId,ColId,Reng,Colum:INteger;
  iTipoHoja:Integer;
  dFecha:TDate;
  sAuxValor:string;
  Encontrado,Salir:Boolean;
  iSaltos:Integer;
  sFrente,IdRecurso:string;
  IidDiario:Integer;
  AuxCol:Integer;
  dCantR,dFactR:Double;
  QrBuscaV,QrBusca,QrVigencia:TZReadOnlyQuery;
  sIdPernocta,sTFiltro,sParamContrato:string;
  QRecurso:TZQuery;
  sHoraInicio,sHoraFinal:string;
  Valido:Boolean;
  dCantHH:Variant;
begin
  {$REGION 'Creacion e Inicializacion de Variables/Objetos'}
  QrBuscaV:=TZReadOnlyQuery.Create(nil);
  QrBuscaV.Connection:=connection.zConnection;

  QrBusca:=TZReadOnlyQuery.Create(nil);
  QrBusca.Connection:=connection.zConnection;

  QrVigencia:=TZReadOnlyQuery.Create(nil);
  QrVigencia.Connection:=connection.zConnection;

  QRecurso:=TZQuery.Create(nil);
  QRecurso.Connection:=connection.zConnection;

  sIdPernocta:='';
  {$ENDREGION}

  try

    if dlgOpenXls.Execute then
    begin
      try
        Excel := CreateOleObject('Excel.Application');
      except
        FreeAndNil(Excel);
        showmessage('No es posible generar el ambiente de EXCEL, informe de esto al administrador del sistema.');
        Exit;
      end;

      Excel.Visible := False;
      Excel.DisplayAlerts := False;
      Excel.ScreenUpdating := False;

      Excel.workbooks.open(dlgOpenXls.FileName);
      RengId:= 5;
      ColId:=9;
      Valido:=True;
    
      try

        for I := 1 to Excel.workbooks[Excel.workbooks.count].WorkSheets.count do
        begin
          Hoja := Excel.WorkBooks[Excel.WorkBooks.Count].WorkSheets[i];
          Hoja.select;
          iTipoHoja:=Hoja.Cells[1, 1].value;
          dFecha:=Hoja.Cells[4,2].value;

//          Excel.ActiveSheet.uNProtect('$SISADM');
          if iTipoHoja=1 then
          begin
            if ValidarCategoria(Excel,Hoja,'Personal',dFecha)then
            begin
              Valido:=False;

            end;

          end;


          if iTipoHoja=2 then
          begin

            if ValidarCategoria(Excel,Hoja,'Equipo',dFecha)then
            begin
              Valido:=False;

            end;
          end;
          Excel.ActiveSheet.Protect('$SISADM', False);

        end;

        if Valido then
        
        for I := 1 to Excel.workbooks[Excel.workbooks.count].WorkSheets.count do
        begin


          Reng:=8;;
          Colum:=1;
          Hoja := Excel.WorkBooks[Excel.WorkBooks.Count].WorkSheets[i];
          Hoja.select;
          iTipoHoja:=Hoja.Cells[1, 1].value;
          dFecha:=Hoja.Cells[4,2].value;
          Salir:=False;

          connection.QryBusca.Active := False;
          connection.QryBusca.SQL.Clear;
          connection.QryBusca.sql.text:='select * from embarcacion_vigencia ' +
                                        'where scontrato=:Contrato and dFechaInicio='+
                                        '(select max(dFechaInicio) from embarcacion_vigencia ' +
                                        'where scontrato=:Contrato and dFechaInicio<=:Fecha)';
          connection.QryBusca.ParamByName('contrato').AsString:=global_Contrato_Barco;
          connection.QryBusca.ParamByName('Fecha').AsDate:=dFecha;
          connection.QryBusca.Open;

          if connection.QryBusca.RecordCount=1 then
            sIdPernocta:=connection.QryBusca.FieldByName('sIdEmbarcacion').AsString;

          if dFecha=tdIdFecha.Date then
          begin
            while not Salir do
            begin
              Encontrado:=False;
              iSaltos:=0;

              if iTipoHoja=1 then
                sTFiltro:=global_labelPersonal;

              if iTipoHoja=2 then
                sTFiltro:=global_labelEquipo;



              while not Encontrado do
              begin
                sAuxValor:=Hoja.Cells[Reng, 1].value;
                if Trim(sAuxValor)<>'' then
                  Encontrado:=True
                else
                begin
                  Inc(iSaltos);
                  Inc(reng);
                end;

                if Encontrado then
                begin
                  QrBusca.Active:=False;
                  QrBusca.SQL.Text:='select c.scontrato from contratos c inner join ordenesdetrabajo ot '+
                                    'on (ot.scontrato=c.scontrato) where c.scodigo=:ContratoBarco and ' +
                                    'ot.snumeroorden=:Orden' ;
                  QrBusca.ParamByName('ContratoBarco').AsString:=global_Contrato_Barco;
                  QrBusca.ParamByName('Orden').AsString:=Trim(sAuxValor);
                  QrBusca.Open;

                  if QrBusca.RecordCount=1 then
                    sParamContrato:=QrBusca.FieldByName('sContrato').AsString;
                end;

                if iSaltos=8 then
                begin
                  Salir:=True;
                  Encontrado:=True;
                end;
              end;

              if not Salir then
              begin
                repeat
                  sFrente:=Hoja.Cells[Reng, 1].value;
                  if Trim(sFrente)<>'' then
                  begin
                    IidDiario:=Hoja.Cells[Reng, 3].value;
                    AuxCol:=ColId;
                    sHoraInicio:=Hoja.Cells[Reng, 6].value;
                    sHoraFinal:=Hoja.Cells[Reng, 7].value;
                    connection.zCommand.Active := False;
                    connection.zCommand.SQL.Clear;
                    connection.zCommand.SQL.Add('Delete from '+ArrayTipo[iTipoHoja,1]{bitacoradepersonal,bitacoradeequipos}+' '+
                                'where sContrato =:contrato and dIdFecha =:fecha and sidPernocta=:pernocta and iidDiario=:Diario');//+ArrayTipo[iTipoHoja,3]{sIdpersonal,sIdequipo}+' =:Personal ');
                    connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
                    connection.zCommand.Params.ParamByName('Contrato').Value    := sParamContrato;
                    connection.zCommand.Params.ParamByName('Fecha').DataType    := ftDate;
                    connection.zCommand.Params.ParamByName('Fecha').Value       := dFecha;
                    connection.zCommand.ParamByName('Diario').AsInteger:=IidDiario;
                    connection.zCommand.Params.ParamByName('Pernocta').AsString    := sIdPernocta;//LoboAzul27
                    connection.zCommand.ExecSQL;

                    repeat
                      sAuxValor:=Hoja.Cells[RengId+3,AuxCol].value;
                      if (Trim(sAuxValor)<>'') then
                      begin
                        IdRecurso:=Hoja.Cells[RengId,AuxCol].value;
                        dCantR:=Hoja.Cells[Reng,AuxCol].value;
                        if dCantR<>0 then
                        begin
                          dCantHH:=Hoja.Cells[Reng,AuxCol+1].value;
                          QrVigencia.Active:=False;
                          QrVigencia.SQL.Text:= 'select mro.*,p.*,ot.sIdPlataforma from movtorecursosxoficio mro inner join ordenesdetrabajogral otg ' +
                                                'on (mro.sContrato=otg.sContrato and mro.sNumeroOrden=otg.sNumeroOrden and mro.dFechaVigencia=otg.dFechaVigencia and mro.IFolioOficio=otg.IFolioOficio) ' +
                                                'inner join '+ArrayTipo[iTipoHoja,2]{personal,equipo}+' p on (p.sContrato = otg.sContrato and p.'+ArrayTipo[iTipoHoja,3]{sIdpersonal,sIdequipo}+' = mro.sNumeroActividad'{sIdpersonal,sIdequipo}+' and p.lCobro ="'+ArrayTipo[iTipoHoja,4]{"Si","No"}+'") '+
                                                'inner join ordenesdetrabajo ot on(ot.scontrato=otg.sContrato and ot.sNumeroOrden=otg.snumeroorden)'+
                                                'where otg.sContrato=:Contrato and otg.sNumeroOrden=:Orden and otg.dFiProgramado=' +
                                                '(select max(otg2.dFiProgramado) from ordenesdetrabajogral otg2 where otg2.sContrato=otg.sContrato and otg2.sNumeroOrden=otg.sNumeroOrden ' +
                                                'and otg2.dFiProgramado <=:Fecha) and mro.sAnexo=:Tipo and mro.snumeroactividad=:Id';
                          QrVigencia.ParamByName('Contrato').AsString:=sParamContrato;
                          QrVigencia.ParamByName('Orden').AsString:=sFrente;
                          QrVigencia.ParamByName('Fecha').AsDate:=dFecha;
                          QrVigencia.ParamByName('Tipo').AsString:=sTFiltro;
                          QrVigencia.ParamByName('Id').AsString:=IdRecurso;
                          QrVigencia.Open;

                          if QrVigencia.RecordCount=1 then
                          begin
                              QRecurso.SQL.Clear;
                              if iTipoHoja <> 2 then
                                 QRecurso.SQL.Add('insert into '+ArrayTipo[iTipoHoja,1]{bitacoradepersonal,bitacoradeequipo}+' (sContrato, dIdFecha, iIdDiario, sIdPersonal, iItemOrden, sDescripcion, sIdPernocta, sIdPlataforma, sHoraInicio, sHoraFinal, dCantidad,dCantHH, sAgrupaPersonal) '+
                                                   'values (:Contrato, :fecha, :diario, :personal, :item, :descripcion, :pernocta, :plataforma, :inicio, :final, :Cantidad,:CantHH, :agrupa)')
                              else
                                 QRecurso.SQL.Add('insert into '+ArrayTipo[iTipoHoja,1]{bitacoradepersonal,bitacoradeequipo}+' (sContrato, dIdFecha, iIdDiario, sIdEquipo, iItemOrden, sDescripcion, sIdPernocta, sHoraInicio, sHoraFinal, dCantidad,dCantHH) '+
                                                   'values (:Contrato, :fecha, :diario, :personal, :item, :descripcion, :pernocta, :inicio, :final, :Cantidad,:CantHH)');
                              QRecurso.ParamByName('Contrato').AsString    := sParamContrato;//global_contrato;
                              QRecurso.ParamByName('fecha').AsDate         := dFecha;
                              QRecurso.ParamByName('diario').AsInteger     := iIdDiario;
                              QRecurso.ParamByName('personal').AsString    := IdRecurso;
                              QRecurso.ParamByName('item').AsInteger       := QrVigencia.FieldByName('iItemOrden').AsInteger;
                              QRecurso.ParamByName('descripcion').AsString := QrVigencia.FieldByName('sDescripcion').AsString;
                              QRecurso.ParamByName('pernocta').AsString    := sIdPernocta;
                              if iTipoHoja <> 2 then
                              begin
                                  QRecurso.ParamByName('plataforma').AsString  := QrVigencia.FieldByName('sIdPlataforma').AsString;
                                  QRecurso.ParamByName('agrupa').AsString      := QrVigencia.FieldByName('sAgrupaPersonal').AsString;
                              end;
                              QRecurso.ParamByName('inicio').AsString      := sHoraInicio;   //rxPersonal.FieldValues['sHoraInicio'+IntToStr(indice)];
                              QRecurso.ParamByName('final').AsString       := sHoraFinal;   //rxPersonal.FieldValues['sHoraFinal'+IntToStr(indice)];
                              QRecurso.ParamByName('cantidad').AsFloat     := dCantR;//rxPersonal.FieldValues['dCantidad'+IntToStr(indice)];
                              QRecurso.ParamByName('CantHH').Value:= dCantHH;
                              QRecurso.ExecSQL;
                          end
                          else
                          begin
                            //NO existe Vigencia P ste Recurso
                          end;

                      
                        end;
                        Inc(AuxCol,2);
                      end;
                    until (Trim(sAuxValor)='') ;
                    Inc(Reng);
                  end;
                until (Trim(sFrente)='');
              end;
            end;
          end;
        end;
      finally
        Excel.ScreenUpdating := True;
        Excel.DisplayAlerts := True;
        Excel.Visible:=true;
      end;
    end;
  finally
    FreeAndNil(QrBuscaV);
    FreeAndNil(QrBusca);

    FreeAndNil(QrVigencia);
    FreeAndNil(QRecurso);
  end;
end;

procedure TFrmCuadreXPartida.btnPernoctaClick(Sender: TObject);
var
  Pos: TPoint;
begin
  if AvSPpmPernoctas.MenuItems.Count>0 then
  begin
    GetCursorPos(Pos);
    AvSPpmPernoctas.ShowMenu((pos.x-btnx) ,(pos.Y + (btnPernocta.Height-btny) ));
  end
  else
    MessageDlg('No hay Pernoctas Disponibles',mtInformation,[mbOK],0);
end;

procedure TFrmCuadreXPartida.btnPernoctaMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
  btnx:=x;
  btny:=y;
end;

procedure TFrmCuadreXPartida.btnPostClick(Sender: TObject);
begin
  GrabarInformacion;
end;

procedure TFrmCuadreXPartida.CargaDatos(sParamContrato,sParamOrden:string;sParamDatos:TsuperPanel);
var
    indice, maximo,Reng,i : integer;
    dCategoria, dTotalDia, dTotalOrden : double;
    QrActividades,QrColumnas,QrDatos,
    QrConcepto,QrVigencia:TZReadOnlyQuery;
   // ColumnTmp:TNxCustomColumn;
    NxTextColumn1,NxTextColumn2: TNxTextColumn;
    NxColPdas:TNxTextColumn;
    ColPda:TPartida;
    RengRecurso:TRecurso;
    sIdPernocta:string;
begin
  {$REGION 'Crear E Inicializar'}
  with sParamDatos do
  begin
    NxGridDatos.BeginUpdate;
    NxGridDatos.Columns.Clear;
    AvGridTurnos.ColCount:=2;
    AvGridPdas.ColCount:=2;
    PSbDatos.Visible:=False;
    ListaPdas.Clear;
    ListaRecursos.Clear;
    NxGridTotales.ClearRows;
  end;

  QrActividades:=TZReadOnlyQuery.Create(nil);
  QrActividades.Connection:=connection.zConnection;

  QrColumnas:=TZReadOnlyQuery.Create(nil);
  QrColumnas.Connection:=connection.zConnection;

  QrDatos:=TZReadOnlyQuery.Create(nil);
  QrDatos.Connection:=connection.zConnection;

  QrConcepto:=TZReadOnlyQuery.Create(nil);
  QrConcepto.Connection:=connection.zConnection;

  QrVigencia:=TZReadOnlyQuery.Create(nil);
  QrVigencia.Connection:=connection.zConnection;

  {$ENDREGION}

  connection.QryBusca.Active := False;
  connection.QryBusca.SQL.Clear;
  connection.QryBusca.sql.text:='select * from embarcacion_vigencia ' +
                                'where scontrato=:Contrato and dFechaInicio='+
                                '(select max(dFechaInicio) from embarcacion_vigencia ' +
                                'where scontrato=:Contrato and dFechaInicio<=:Fecha)';
  connection.QryBusca.ParamByName('contrato').AsString:=global_Contrato_Barco;
  connection.QryBusca.ParamByName('Fecha').AsDate:=tdIdFecha.Date;
  connection.QryBusca.Open;

  if connection.QryBusca.RecordCount=1 then
    sIdPernocta:=connection.QryBusca.FieldByName('sIdEmbarcacion').AsString;


  QrActividades.SQL.Text:='select * from bitacoradeactividades '+
                          'where sContrato =:Contrato and sIdTurno =:Turno and sNumeroOrden =:Orden and dIdFecha =:fecha '+
                          'and sIdTipoMovimiento ="'+ArrayTipo[iTipo,5]{"E","N"}+'" order by sHoraInicio ';
  QrActividades.ParamByName('Contrato').AsString := sParamContrato;
  QrActividades.ParamByName('Fecha').AsDate      := tdIdFecha.Date;
  QrActividades.ParamByName('Orden').AsString    := sParamOrden;
  QrActividades.ParamByName('Turno').AsString    := global_turno;
  QrActividades.Open;

  QrColumnas.SQL.Text:='select a.sContrato, a.iIdDiario, a.sNumeroActividad,concat(a.sHoraInicio,"-",a.sHorafinal)  as Horario, tm.sIdTipoMovimiento,a.sIdClasificacion,a.sHoraInicio,a.sHorafinal from bitacoradeactividades a ';
  if iTipo <> 3 then
    QrColumnas.SQL.add('inner join actividadesxorden o on (o.sContrato = a.sContrato and o.sIdConvenio =:convenio and o.sNumeroOrden =:orden and a.sWbs = o.sWbs and o.sNumeroActividad = a.sNumeroActividad and o.sTipoActividad = "Actividad") ')
  else
    QrColumnas.SQL.Add('left join actividadesxorden o on (o.sContrato = a.sContrato and o.sIdConvenio =:convenio and o.sNumeroOrden =:orden and a.sWbs = o.sWbs and o.sNumeroActividad = a.sNumeroActividad and o.sTipoActividad = "Actividad") ');

  QrColumnas.SQL.Add( 'inner join tiposdemovimiento tm on(tm.sContrato=:ContratoBarco and tm.sIdTipoMovimiento=a.sIdTipoMovimiento and tm.sClasificacion="Tarifa Diaria")');
  QrColumnas.SQL.Add( 'where a.sContrato =:contrato and a.dIdFecha =:fecha and a.sNumeroOrden =:orden and a.sIdTurno =:turno and a.sIdTipoMovimiento = "'+ArrayTipo[iTipo,5]{"E","N"}+'" '+
                      'group by a.sContrato, a.iIdDiario '+
                      'order by a.sContrato,o.iItemOrden,a.sHoraInicio');
  QrColumnas.ParamByName('Contrato').AsString := sParamContrato;
  QrColumnas.ParamByName('ContratoBarco').AsString := Global_contrato_Barco;
  QrColumnas.ParamByName('Convenio').AsString := global_convenio;
  QrColumnas.ParamByName('Orden').AsString    := sParamOrden;
  QrColumnas.ParamByName('Turno').AsString    := global_turno;
  QrColumnas.ParamByName('fecha').AsDate      := tdIdFecha.Date;
  QrColumnas.Open;

  with sParamDatos do
  begin
    AvGridTurnos.ColWidths[0]:=12;
    AvGridTurnos.ColWidths[1]:=320;
    AvGridTurnos.Cols[1].Add('PARTIDA');
    AvGridPdas.ColWidths[0]:=12;
    AvGridPdas.ColWidths[1]:=320;
    AvGridPdas.Cols[1].Add('CLASIFICACION');

    NxTextColumn1 := TNxTextColumn.Create(NxGridDatos);
    NxTextColumn1.Header.Caption:= 'ID';
    NxTextColumn1.Header.Alignment:= Talignment(2);
    NxTextColumn1.Options := [coDontHighlight, coPublicUsing, coShowTextFitHint,coDisableMoving];
    NxTextColumn1.Position := 0;
    NxTextColumn1.DefaultWidth := 70;
    NxTextColumn1.Width := 70;
    NxTextColumn1.Alignment := Talignment(3);
    NxTextColumn1.Color:=$00D8D8D8;
    NxTextColumn1.Font.Style:=[fsBold];
    NxTextColumn1.Font.Color:=clNavy;
    NxGridDatos.Columns.AddColumn(NxTextColumn1);

    NxTextColumn2 := TNxTextColumn.Create(NxGridDatos);
    NxTextColumn2.Header.Caption:= 'DESCRIPCION';
    NxTextColumn2.Header.Alignment:= Talignment(2);
    NxTextColumn2.Options := [coDontHighlight, coPublicUsing, coShowTextFitHint,coDisableMoving];
    NxTextColumn2.Position := 0;
    NxTextColumn2.DefaultWidth := 250;
    NxTextColumn2.Width := 250;
    NxTextColumn2.Alignment := Talignment(3);
    NxTextColumn2.Color:=$00D8D8D8;
    NxTextColumn2.Font.Style:=[fsBold];
    NxGridDatos.Columns.AddColumn(NxTextColumn2);
  end;

  indice:=2;
  while not QrColumnas.Eof do
  begin
    NxColPdas:=TNxTextColumn.Create(sParamDatos.NxGridDatos);
    NxColPdas.Name:='Pdas' + IntToStr(Indice);
    NxColPdas.Header.Caption:=QrColumnas.FieldByName('Horario').AsString;
    NxColPdas.Header.Alignment:= Talignment(2);
    NxColPdas.Alignment := Talignment(2);
    NxColPdas.Position := 0;
    NxColPdas.DefaultWidth :=80;
    NxColPdas.Width := 80;
    NxColPdas.Options := [coCanClick,coPublicUsing,coCanInput,coEditing,coDisableMoving];
    sParamDatos.NxGridDatos.Columns.AddColumn(NxColPdas);

    with sParamDatos do
    begin
      AvGridTurnos.AddColumn;
      AvGridTurnos.Cols[indice].Add(QrColumnas.FieldByName('sNumeroActividad').AsString);
      AvGridTurnos.ColWidths[indice]:=80;

      AvGridPdas.AddColumn;
      AvGridPdas.Cols[indice].Add(QrColumnas.FieldByName('sIdClasificacion').AsString);
      AvGridPdas.ColWidths[indice]:=80;
    end;

    ColPda:=TPartida.Create;
    with ColPda,QrColumnas do
    begin
      IdDiario:=FieldByName('iidDiario').AsInteger;
      iCol:=indice;
      sPartida:=FieldByName('snumeroactividad').AsString;
      sIdClasificacion:=FieldByName('sIdClasificacion').AsString;
      sHoraInicio:=FieldByName('sHoraInicio').AsString;
      sHoraFinal:=FieldByName('sHoraFinal').AsString;
    end;
    sParamDatos.ListaPdas.AddObject(IntToStr(indice),ColPda);
    Inc(Indice);
    QrColumnas.Next;
  end;

  with sParamDatos do
  begin
    if AvGridPdas.ColCount>2 then
    begin
     AvGridPdas.FixedCols:=2;
     AvGridTurnos.FixedCols:=2;
    end;

    if NxGridDatos.Columns.Count>2 then
      NxGridDatos.FixedCols:=2;

    if NxGridDatos.Columns.Count>13 then
    begin

      PSbDatos.Visible:=True;
      AvGridPdas.col:=12;
      AvGridTurnos.col:=12;
      PSbDatos.SetParams(0,0,(NxGridDatos.Columns.Count - 13));
      PSbDatos.LargeChange:=1;
      PSbDatos.SmallChange:=1;
    end;
  end;

  //Aqui van las vigencias de personal
  QrVigencia.Active:=False;
  QrVigencia.SQL.Text:= 'select mro.*,p.*,ot.sIdPlataforma from movtorecursosxoficio mro inner join ordenesdetrabajogral otg ' +
                        'on (mro.sContrato=otg.sContrato and mro.sNumeroOrden=otg.sNumeroOrden and mro.dFechaVigencia=otg.dFechaVigencia and mro.IFolioOficio=otg.IFolioOficio) ' +
                        'inner join '+ArrayTipo[iTipo,2]{personal,equipo}+' p on (p.sContrato = otg.sContrato and p.'+ArrayTipo[iTipo,3]{sIdpersonal,sIdequipo}+' = mro.sNumeroActividad'{sIdpersonal,sIdequipo}+' and p.lCobro ="'+ArrayTipo[iTipo,4]{"Si","No"}+'") '+
                        'inner join ordenesdetrabajo ot on(ot.scontrato=otg.sContrato and ot.sNumeroOrden=otg.snumeroorden)'+
                        'where otg.sContrato=:Contrato and otg.sNumeroOrden=:Orden and otg.dFiProgramado=' +
                        '(select max(otg2.dFiProgramado) from ordenesdetrabajogral otg2 where otg2.sContrato=otg.sContrato and otg2.sNumeroOrden=otg.sNumeroOrden ' +
                        'and otg2.dFiProgramado <=:Fecha) and mro.sAnexo=:Tipo';
  QrVigencia.ParamByName('Contrato').AsString:=sParamContrato;
  QrVigencia.ParamByName('Orden').AsString:=sParamOrden;
  QrVigencia.ParamByName('Fecha').AsDate:=tdIdFecha.Date;
  QrVigencia.ParamByName('Tipo').AsString:=sTipoFiltro;
  QrVigencia.Open;



  {'select mr.* from moerecursos mr inner join moe m on (m.iidMoe=mr.iidMoe) ' +
                        'where m.scontrato=:Contrato and didfecha =(select max(didfecha) from moe where didfecha <=:Fecha)';}

  QrConcepto.SQL.Add( 'select b.* from '+ArrayTipo[iTipo,1]{bitacoradepersonal,bitacoradeequipos}+' b '+
                      'inner join '+ArrayTipo[iTipo,2]{personal,equipo}+' p on (p.sContrato = b.sContrato and p.'+ArrayTipo[iTipo,3]{sIdpersonal,sIdequipo}+' = b.'+ArrayTipo[iTipo,3]{sIdpersonal,sIdequipo}+' and p.lCobro ="'+ArrayTipo[iTipo,4]{"Si","No"}+'") '+
                      'where b.sContrato =:contrato and b.dIdFecha =:fecha and b.sidpernocta=:Pernocta ' +
                      'group by b.'+ArrayTipo[iTipo,3]{sIdpersonal,sIdequipo}+' '+
                      'order by b.iItemOrden, b.'+ArrayTipo[iTipo,3]{sIdpersonal,sIdequipo}+'');
  QrConcepto.ParamByName('Contrato').AsString := sParamContrato;
  QrConcepto.ParamByName('fecha').AsDate      := tdIdFecha.Date;
  QrConcepto.ParamByName('Pernocta').AsString:=sIdPernocta;//LoboAzul27
  QrConcepto.Open;

  while not QrVigencia.Eof do
  begin
    with sParamDatos do
    begin
      NxGridDatos.AddRow();
      NxGridDatos.Cell[0,NxGridDatos.LastAddedRow].AsString:=QrVigencia.fieldByname('sNumeroActividad').AsString;  //QrConcepto.Fieldbyname(ArrayTipo[iTipo,3]).AsString;
      NxGridDatos.Cell[1,NxGridDatos.LastAddedRow].AsString:=QrVigencia.FieldByName('sDescripcion').AsString;

    end;

    RengRecurso:=TRecurso.Create;
    RengRecurso.sIdRecurso:=QrVigencia.fieldByname('sNumeroActividad').AsString;
    RengRecurso.sDescripcion:=QrVigencia.FieldByName('sDescripcion').AsString;
    RengRecurso.ItemOrden:=QrVigencia.FieldByName('iItemOrden').AsInteger;
    RengRecurso.iReng:=sParamDatos.NxGridDatos.LastAddedRow;
    if iTipo <> 2 then
    begin
      RengRecurso.sAgrupa:=QrVigencia.FieldByName('sAgrupaPersonal').AsString;
      RengRecurso.sPlataforma:=QrVigencia.FieldByName('sIdPlataforma').AsString;
    end;
    sParamDatos.ListaRecursos.AddObject(RengRecurso.sIdRecurso,RengRecurso);
    
    QrVigencia.Next;
  end;

  QrDatos.SQL.Add('select  a.sNumeroActividad,sum(b.dCantidad) as dCantidad, b.* from '+ArrayTipo[iTipo,1]{bitacoradepersonal,bitacoradeequipos}+' b '+
                  'inner join bitacoradeactividades a on (a.sContrato = b.sContrato and a.dIdFecha = b.dIdFecha and b.iIdDiario = a.iIdDiario and a.sNumeroOrden =:orden and a.sIdTurno =:turno and a.sIdTipoMovimiento = "'+ArrayTipo[iTipo,5]{"E","N"}+'") ');

  if iTipo <> 3 then
    QrDatos.SQL.Add('inner join actividadesxorden o on (o.sContrato = a.sContrato and o.sIdConvenio =:convenio and o.sNumeroOrden =a.sNumeroOrden and a.sWbs = o.sWbs and o.sNumeroActividad = a.sNumeroActividad and o.sTipoActividad = "Actividad") ')
  else
    QrDatos.SQL.Add('left join actividadesxorden o on (o.sContrato = a.sContrato and o.sIdConvenio =:convenio and o.sNumeroOrden =a.sNumeroOrden and a.sWbs = o.sWbs and o.sNumeroActividad = a.sNumeroActividad and o.sTipoActividad = "Actividad") ');

  QrDatos.SQL.Add(
                  'inner join '+ArrayTipo[iTipo,2]{personal,equipo}+' p on (p.sContrato = b.sContrato and p.'+ArrayTipo[iTipo,3]{sIdPersonal,sIdEquipo}+' = b.'+ArrayTipo[iTipo,3]{sIsPersonal,sIdEquipo}+' and p.lCobro = "'+ArrayTipo[iTipo,4]{"Si","No"}+'") '+
                  'where b.sContrato =:Contrato and b.dIdFecha =:fecha and b.sidpernocta=:Pernocta '+
                  'group by b.sContrato, b.'+ArrayTipo[iTipo,3]{sIdPersonal,sIdEquipo}+',b.iIdDiario '+
                  'order by b.sContrato, o.iItemOrden');
  QrDatos.ParamByName('Contrato').AsString := sParamContrato;
  QrDatos.ParamByName('Convenio').AsString := global_convenio;
  QrDatos.ParamByName('Orden').AsString    := sParamOrden;
  QrDatos.ParamByName('Turno').AsString    := global_turno;
  QrDatos.ParamByName('fecha').AsDate      := tdIdFecha.Date;
  QrDatos.ParamByName('Pernocta').AsString:=sIdPernocta;//QrPernoctas.FieldByName('sIdPernocta').AsString;//LoboAzul27
  QrDatos.Open;

  Reng:=0;
  with sParamDatos do
  begin
    while Reng<NxGridDatos.RowCount do
    begin
      QrDatos.Filtered:=False;
      QrDatos.Filter:=ArrayTipo[iTipo,3]+'='+quotedstr(NxGridDatos.Cell[0,Reng].AsString);
      QrDatos.Filtered:=True;

      QrDatos.First;
      while not QrDatos.Eof do
      begin
        for I := 0 to ListaPdas.Count - 1 do
          if TPartida(ListaPdas.Objects[I]).IdDiario=QrDatos.FieldByName('iIdDiario').AsInteger then
          begin
            NxGridDatos.Cell[TPartida(ListaPdas.Objects[I]).iCol,reng].AsInteger:=QrDatos.FieldByName('dCantidad').AsInteger;
            Break;
          end;
        QrDatos.Next;
      end;

      dTotalDia:=0;
      for I := 2 to NxGridDatos.Columns.Count - 1 do
        dTotalDia:=dTotalDia + StrToFloatDef(NxGridDatos.Cell[i,reng].AsString,0);

      NxGridTotales.AddRow();
      NxGridTotales.Cell[0,NxGridTotales.LastAddedRow].Asfloat:=dTotalDia;

      Inc(Reng);
    end;

    if NxGridDatos.Columns.Count>13 then
    begin
      NxGridDatos.HorzScrollBar.AutoHide:=False;
      NxGridDatos.HorzScrollBar.Visible:=False;
    end;
    NxGridDatos.EndUpdate;
    
    if NxGridDatos.RowCount>24 then
    begin
      NxGridDatos.VertScrollBar.AutoHide:=False;
      NxGridDatos.VertScrollBar.Visible:=False;
      NxGridDatos.MouseWheelEnabled:=true;
    end
    else
      NxGridDatos.MouseWheelEnabled:=False;
  end;
  
end;


procedure TFrmCuadreXPartida.cmdAceptarClick(Sender: TObject);
var
  RengRecurso:TRecurso;
  sDatosPanel:TSuperPanel;
begin
  sDatosPanel:=TSuperPanel(ListaObj.Objects[ListaObj.IndexOf(NxPgcDatos.Pages[NxPgcDatos.ActivePageIndex].Name)]);
  if sDatosPanel.ListaRecursos.IndexOf(catalogo_maestro.FieldByName(dbUgridBuscar.Columns[0].FieldName).AsString)=-1 then
  begin
    with sDatosPanel,catalogo_maestro do
    begin
      NxGridDatos.BeginUpdate;
      NxGridDatos.AddRow();
      NxGridDatos.Cell[0,NxGridDatos.LastAddedRow].AsString:=Fieldbyname(dbUgridBuscar.Columns[0].FieldName).AsString;
      NxGridDatos.Cell[1,NxGridDatos.LastAddedRow].AsString:=FieldByName('sDescripcion').AsString;

       //qry_catalogo.FieldValues[ArrayTipo[iTipo,3]
      RengRecurso:=TRecurso.Create;
      RengRecurso.sIdRecurso:=Fieldbyname(dbUgridBuscar.Columns[0].FieldName).AsString;
      RengRecurso.sDescripcion:=FieldByName('sDescripcion').AsString;
      RengRecurso.ItemOrden:=FieldByName('iItemOrden').AsInteger;
      RengRecurso.iReng:=NxGridDatos.LastAddedRow;
      if iTipo <> 2 then
      begin
        RengRecurso.sAgrupa:=FieldByName('sAgrupaPersonal').AsString;
        RengRecurso.sPlataforma:=ordenesdetrabajo.FieldByName('sIdPlataforma').AsString;
      end;
      ListaRecursos.AddObject(RengRecurso.sIdRecurso,RengRecurso);
      NxGridTotales.AddRow();
      NxGridTotales.Cell[0,NxGridTotales.LastAddedRow].Asfloat:=0;

      if NxGridDatos.Columns.Count>13 then
      begin
        NxGridDatos.HorzScrollBar.AutoHide:=False;
        NxGridDatos.HorzScrollBar.Visible:=False;
      end;
      NxGridDatos.EndUpdate;

      if NxGridDatos.RowCount>24 then
      begin
        NxGridDatos.VertScrollBar.AutoHide:=False;
        NxGridDatos.VertScrollBar.Visible:=False;
        NxGridDatos.HorzScrollBar.AutoHide:=False;
        NxGridDatos.HorzScrollBar.Visible:=False;
        NxGridDatos.MouseWheelEnabled:=true;
      end
      else
        NxGridDatos.MouseWheelEnabled:=False;
    end;
  end;
  NGbxConsulta.Visible := False;
  catalogo_maestro.Active := False;

  //sDatosPanel.AvGridTurnos.ScrollBars.
end;

procedure TFrmCuadreXPartida.cmdCancelarClick(Sender: TObject);
begin
  NGbxConsulta.Visible := False;
  catalogo_maestro.Active := False;
end;

procedure TFrmCuadreXPartida.DeUndiaAnterior1Click(Sender: TObject);
var
  sDatosPanel:TSuperPanel;
  iPosDato,I:Integer;
  Recurso:TRecurso;
  dFechaAnterior : tDate;
  lNuevaVigencia : boolean;
  lEncuentra     : boolean;
  iResp:Integer;
  QrConcepto:TZReadOnlyQuery;
  RengRecurso:TRecurso;
  dTotalDia:Double;
begin
  sDatosPanel:=TSuperPanel(ListaObj.Objects[ListaObj.IndexOf(NxPgcDatos.Pages[NxPgcDatos.ActivePageIndex].Name)]);
  if sDatosPanel.NxGridDatos.Columns.Count<3 then
  begin
    MessageDlg('No existen partidas Reportadas en el Día!',mtConfirmation, [mbYes, mbNo], 0);
    exit;
  end;
  dFechaAnterior := tdIdFecha.Date - 1;
  QrConcepto:=TZReadOnlyQuery.Create(nil);
  QrConcepto.Connection:=connection.zConnection;

(*  Connection.qryBusca.SQL.Clear;
  Connection.qryBusca.SQL.Add('Select * from '+ArrayTipo[iTipo,1]{bitacoradepersonal,bitacoradeequipos}+' Where sContrato =:Contrato and dIdFecha =:Fecha ' +
                               'and sidpernocta=:Pernocta');
  Connection.qryBusca.Params.ParamByName('Contrato').Datatype := ftString;
  Connection.qryBusca.Params.ParamByName('Contrato').Value    := global_contrato;
  Connection.qryBusca.Params.ParamByName('Fecha').Datatype    := ftDate;
  Connection.qryBusca.Params.ParamByName('Fecha').Value       := tdIdFecha.Date;
  Connection.qryBusca.ParamByName('pernocta').AsString:=sDatosPanel.Id;  //DblCmbPernoctas.KeyValue;
  Connection.qryBusca.Open; *)
  QrConcepto.SQL.Add( 'select b.* from '+ArrayTipo[iTipo,1]{bitacoradepersonal,bitacoradeequipos}+' b '+
                      'inner join '+ArrayTipo[iTipo,2]{personal,equipo}+' p on (p.sContrato = b.sContrato and p.'+ArrayTipo[iTipo,3]{sIdpersonal,sIdequipo}+' = b.'+ArrayTipo[iTipo,3]{sIdpersonal,sIdequipo}+' and p.lCobro ="'+ArrayTipo[iTipo,4]{"Si","No"}+'") '+
                      'where b.sContrato =:contrato and b.dIdFecha =:fecha and b.sidpernocta=:Pernocta ' +
                      'group by b.'+ArrayTipo[iTipo,3]{sIdpersonal,sIdequipo}+' '+
                      'order by b.iItemOrden, b.'+ArrayTipo[iTipo,3]{sIdpersonal,sIdequipo}+'');
  QrConcepto.ParamByName('Contrato').AsString := global_contrato;
  QrConcepto.ParamByName('fecha').AsDate      := dFechaAnterior;
  QrConcepto.ParamByName('Pernocta').AsString:=sDatosPanel.Id;
  QrConcepto.Open;

  if QrConcepto.RecordCount>0 then
  begin
    sDatosPanel.NxGridDatos.BeginUpdate;
    iResp:=mrYes;
    if sDatosPanel.NxGridDatos.RowCount > 0 then
    begin
      iResp:=MessageDlg('Desea Eliminar el '+ArrayTipo[iTipo,2]{personal,equipo}+' del Día Actual?',mtConfirmation, [mbYes, mbNo,mbCancel], 0);
      if iResp= mrYes then
      begin
        Connection.qryBusca2.SQL.Clear;
        Connection.qryBusca2.SQL.Add('Select '+ArrayTipo[iTipo,3]{sIdpersonal,sIdequipo}+' from '+ArrayTipo[iTipo,2]{personal,equipo}+' Where sContrato =:Contrato and lCobro ="'+ArrayTipo[iTipo,4]{"Si","No"}+'" ');
        Connection.qryBusca2.Params.ParamByName('Contrato').Datatype := ftString;
        Connection.qryBusca2.Params.ParamByName('Contrato').Value    := global_contrato;

        Connection.qryBusca2.Open;

        while not connection.QryBusca2.Eof do
        begin
          connection.zCommand.Active := False;
          connection.zCommand.SQL.Clear;
          connection.zCommand.SQL.Add('Delete from '+ArrayTipo[iTipo,1]{bitacoradepersonal,bitacoradeequipos}+' '+
                    'where sContrato =:contrato and dIdFecha =:fecha and '+ArrayTipo[iTipo,3]{sIdpersonal,sIdequipo}+' =:Personal '+
                      'and sidpernocta=:Pernocta');
          connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
          connection.zCommand.Params.ParamByName('Contrato').Value    := global_contrato;
          connection.zCommand.Params.ParamByName('Fecha').DataType    := ftDate;
          connection.zCommand.Params.ParamByName('Fecha').Value       := tdIdFecha.Date;
          connection.zCommand.Params.ParamByName('Personal').DataType := ftString;
          connection.zCommand.Params.ParamByName('Personal').Value    := Connection.qryBusca2.FieldValues[ArrayTipo[iTipo,3]{sIdpersonal,sIdequipo}];
          Connection.zCommand.ParamByName('pernocta').AsString:=sDatosPanel.Id;
          connection.zCommand.ExecSQL;
          connection.QryBusca2.Next;
        end;

        sDatosPanel.NxGridDatos.ClearRows;
        sDatosPanel.NxGridTotales.ClearRows;
        sDatosPanel.ListaRecursos.Clear;
      end;
    end;

    if iResp<> mrCancel
      (*MessageDlg('Desea cargar el '+ArrayTipo[iTipo,2]{personal,equipo}+' del dia anterior?', mtConfirmation, [mbYes, mbNo], 0) = mrYes*)
    then
    begin
      with sDatosPanel do
      begin
        while not QrConcepto.Eof do
        begin
          if sDatosPanel.ListaRecursos.IndexOf(QrConcepto.FieldByName(ArrayTipo[iTipo,3]).AsString)=-1 then
          begin
            NxGridDatos.AddRow();
            NxGridDatos.Cell[0,NxGridDatos.LastAddedRow].AsString:=QrConcepto.Fieldbyname(ArrayTipo[iTipo,3]).AsString;
            NxGridDatos.Cell[1,NxGridDatos.LastAddedRow].AsString:=QrConcepto.FieldByName('sDescripcion').AsString;

             //qry_catalogo.FieldValues[ArrayTipo[iTipo,3]
            RengRecurso:=TRecurso.Create;
            RengRecurso.sIdRecurso:=QrConcepto.Fieldbyname(ArrayTipo[iTipo,3]).AsString;
            RengRecurso.sDescripcion:=QrConcepto.FieldByName('sDescripcion').AsString;
            RengRecurso.ItemOrden:=QrConcepto.FieldByName('iItemOrden').AsInteger;
            RengRecurso.iReng:=NxGridDatos.LastAddedRow;
            if iTipo <> 2 then
            begin
              RengRecurso.sAgrupa:=QrConcepto.FieldByName('sAgrupaPersonal').AsString;
              RengRecurso.sPlataforma:=QrConcepto.FieldByName('sIdPlataforma').AsString;
            end;
            ListaRecursos.AddObject(RengRecurso.sIdRecurso,RengRecurso);
          end;

          Connection.Auxiliar.Active := False;
          Connection.Auxiliar.SQL.Clear;
          Connection.Auxiliar.SQL.Add('select a.sNumeroActividad, t.sHoraInicio, t.sHoraFinal, b.*, a.sIdTurnoHora, a.sNumeroOrden from '+ArrayTipo[iTipo,1]{bitacoradepersonal,bitacoradeequipos}+' b '+
                            'inner join bitacoradeactividades a on (a.sContrato = b.sContrato and a.dIdFecha = b.dIdFecha and b.iIdDiario = a.iIdDiario and a.sNumeroOrden =:orden and a.sIdTurno =:turno and a.sIdTipoMovimiento= "'+ArrayTipo[iTipo,5]{"E","N"}+'") ');
          if iTipo <> 3 then
             Connection.Auxiliar.SQL.Add('inner join actividadesxorden o on (o.sContrato = a.sContrato and o.sIdConvenio =:convenio and o.sNumeroOrden =:orden and a.sWbs = o.sWbs and o.sNumeroActividad = a.sNumeroActividad and o.sTipoActividad = "Actividad") ')
          else
             Connection.Auxiliar.SQL.Add('left join actividadesxorden o on (o.sContrato = a.sContrato and o.sIdConvenio =:convenio and o.sNumeroOrden =:orden and a.sWbs = o.sWbs and o.sNumeroActividad = a.sNumeroActividad and o.sTipoActividad = "Actividad") ');
          Connection.Auxiliar.SQL.Add('inner join turnos_horas t on (t.sContrato = b.sContrato and t.sIdturnoHora = a.sIdTurnoHora and :fecha >= t.dFecha) '+
                            'inner join '+ArrayTipo[iTipo,2]{personal,equipo}+' p on (p.sContrato = b.sContrato and p.'+ArrayTipo[iTipo,3]{sIdpersonal,sIdequipo}+' = b.'+ArrayTipo[iTipo,3]{sIdpersonal,sIdequipo}+' and p.lCobro = "'+ArrayTipo[iTipo,4]{"Si","No"}+'") '+
                            'where b.sContrato =:Contrato and b.dIdFecha =:fecha and b.sidpernocta=:Pernocta and b.'+ArrayTipo[iTipo,3] + '=:Id ' +
                            'group by b.sContrato, b.'+ArrayTipo[iTipo,3]{sIdpersonal,sIdequipo}+', a.sIdturnoHora, b.iIdDiario '+
                            'order by b.sContrato, a.sIdTurnoHora ASC, o.iItemOrden');
          Connection.Auxiliar.Params.ParamByName('Contrato').Datatype := ftString;
          Connection.Auxiliar.Params.ParamByName('Contrato').Value    := global_contrato;
          Connection.Auxiliar.Params.ParamByName('Convenio').Datatype := ftString;
          Connection.Auxiliar.Params.ParamByName('Convenio').Value    := global_convenio;
          Connection.Auxiliar.Params.ParamByName('Turno').Datatype    := ftString;
          Connection.Auxiliar.Params.ParamByName('Turno').Value       := global_turno;
          Connection.Auxiliar.Params.ParamByName('Orden').Datatype    := ftString;
        //  Connection.Auxiliar.Params.ParamByName('Orden').Value       := tsNumeroOrden.KeyValue;
          Connection.Auxiliar.Params.ParamByName('Fecha').Datatype    := ftDate;
          Connection.Auxiliar.Params.ParamByName('Fecha').Value       := dFechaAnterior;
          Connection.Auxiliar.ParamByName('Id').AsString:=QrConcepto.Fieldbyname(ArrayTipo[iTipo,3]).AsString;
          Connection.Auxiliar.ParamByName('pernocta').AsString:=sDatosPanel.Id;
          Connection.Auxiliar.Open;

          connection.zCommand.Active := False;
          connection.zCommand.SQL.Clear;
          if iTipo <> 2 then
            Connection.zcommand.SQL.Add('insert into '+ArrayTipo[iTipo,1]{bitacoradepersonal,bitacoradeequipo}+' (sContrato, dIdFecha, iIdDiario, sIdPersonal, iItemOrden, sDescripcion, sIdPernocta, sIdPlataforma, sHoraInicio, sHoraFinal, dCantidad, sAgrupaPersonal) '+
                                      'values (:Contrato, :fecha, :diario, :personal, :item, :descripcion, :pernocta, :plataforma, :inicio, :final, :Cantidad, :agrupa)')
          else
            Connection.zcommand.SQL.Add('insert into '+ArrayTipo[iTipo,1]{bitacoradepersonal,bitacoradeequipo}+' (sContrato, dIdFecha, iIdDiario, sIdEquipo, iItemOrden, sDescripcion, sIdPernocta, sHoraInicio, sHoraFinal, dCantidad) '+
                                      'values (:Contrato, :fecha, :diario, :personal, :item, :descripcion, :pernocta, :inicio, :final, :Cantidad)');

          lEncuentra := False;
          while not Connection.Auxiliar.Eof do
          begin
            for iPosDato := 0 to ListaPdas.Count-1 do
            begin
              if {( TPartida(ListaPdas.Objects[iPosDato]).sIdTurno = Connection.Auxiliar.FieldValues['sIdTurnoHora']) and}
                (TPartida(ListaPdas.Objects[iPosDato]).sPartida = Connection.Auxiliar.FieldValues['sNumeroActividad']) then
              begin

                NxGridDatos.Cell[TPartida(ListaPdas.Objects[iPosDato]).iCol,
                TRecurso(ListaRecursos.Objects[sDatosPanel.ListaRecursos.IndexOf(QrConcepto.FieldByName(ArrayTipo[iTipo,3]).AsString)]).iReng
                ].AsInteger:=Connection.Auxiliar.FieldByName('dCantidad').AsInteger;

                lEncuentra := True;
                try
                    Connection.zcommand.ParamByName('Contrato').AsString    := Connection.Auxiliar.FieldValues['sContrato'];
                    Connection.zcommand.ParamByName('fecha').AsDate         := tdIdFecha.Date;
                    Connection.zcommand.ParamByName('Diario').AsInteger     := TPartida(ListaPdas.Objects[iPosDato]).IdDiario;
                    Connection.zcommand.ParamByName('personal').AsString    := Connection.Auxiliar.FieldValues[ArrayTipo[iTipo,3]{sIdPersonal,sIdEquipo}];
                    Connection.zcommand.ParamByName('item').AsInteger       := Connection.Auxiliar.FieldValues['iItemOrden'];
                    Connection.zcommand.ParamByName('descripcion').AsString := Connection.Auxiliar.FieldValues['sDescripcion'];
                    Connection.zcommand.ParamByName('pernocta').AsString    := Connection.Auxiliar.FieldValues['sIdPernocta'];
                    if iTipo <> 2 then
                    begin
                        Connection.zcommand.ParamByName('plataforma').AsString  := Connection.Auxiliar.FieldValues['sIdPlataforma'];
                        Connection.zcommand.ParamByName('agrupa').AsString      := Connection.Auxiliar.FieldValues['sAgrupaPersonal'];
                    end;
                    Connection.zcommand.ParamByName('inicio').AsString      := Connection.Auxiliar.FieldValues['sHoraInicio'];
                    Connection.zcommand.ParamByName('final').AsString       := Connection.Auxiliar.FieldValues['sHoraFinal'];
                    Connection.zcommand.ParamByName('cantidad').AsFloat     := Connection.Auxiliar.FieldValues['dCantidad'];
                    Connection.zcommand.ExecSQL;
                except
                end;

                Break;
              end;
            end;
            Connection.Auxiliar.Next;
          end;


          if lEncuentra = False then
          begin
              //Insertamos al menos las categorias en cero a cualquier partida..
              connection.zCommand.Active := False;
              connection.zCommand.SQL.Clear;
              if iTipo <> 2 then
                 Connection.zcommand.SQL.Add('insert into '+ArrayTipo[iTipo,1]{bitacoradepersonal,bitacoradeequipo}+' (sContrato, dIdFecha, iIdDiario, sIdPersonal, iItemOrden, sDescripcion, sIdPernocta, sIdPlataforma, sHoraInicio, sHoraFinal, dCantidad, sAgrupaPersonal) '+
                                          'values (:Contrato, :fecha, :diario, :personal, :item, :descripcion, :pernocta, :plataforma, :inicio, :final, :Cantidad, :agrupa)')
              else
                   Connection.zcommand.SQL.Add('insert into '+ArrayTipo[iTipo,1]{bitacoradepersonal,bitacoradeequipo}+' (sContrato, dIdFecha, iIdDiario, sIdEquipo, iItemOrden, sDescripcion, sIdPernocta, sHoraInicio, sHoraFinal, dCantidad) '+
                                          'values (:Contrato, :fecha, :diario, :personal, :item, :descripcion, :pernocta, :inicio, :final, :Cantidad)');
              try
                  Connection.zcommand.ParamByName('Contrato').AsString    := Connection.Auxiliar.FieldValues['sContrato'];
                  Connection.zcommand.ParamByName('fecha').AsDate         := tdIdFecha.Date;
                  Connection.zcommand.ParamByName('Diario').AsInteger     :=  TPartida(ListaPdas.Objects[0]).IdDiario;
                  Connection.zcommand.ParamByName('personal').AsString    := Connection.Auxiliar.FieldValues[ArrayTipo[iTipo,3]{sIdPpersonal,sIdequipo}];
                  Connection.zcommand.ParamByName('item').AsInteger       := Connection.Auxiliar.FieldValues['iItemOrden'];
                  Connection.zcommand.ParamByName('descripcion').AsString := Connection.Auxiliar.FieldValues['sDescripcion'];
                  Connection.zcommand.ParamByName('pernocta').AsString    := Connection.Auxiliar.FieldValues['sIdPernocta'];
                  if iTipo <> 2 then
                  begin
                      Connection.zcommand.ParamByName('plataforma').AsString  := Connection.Auxiliar.FieldValues['sIdPlataforma'];
                      Connection.zcommand.ParamByName('agrupa').AsString      := Connection.Auxiliar.FieldValues['sAgrupaPersonal'];
                  end;
                  Connection.zcommand.ParamByName('inicio').AsString      := Connection.Auxiliar.FieldValues['sHoraInicio'];
                  Connection.zcommand.ParamByName('final').AsString       := Connection.Auxiliar.FieldValues['sHoraFinal'];
                  Connection.zcommand.ParamByName('cantidad').AsFloat     := 0;
                  Connection.zcommand.ExecSQL;
              except
              end;
          end;

          dTotalDia:=0;
          for I := 2 to NxGridDatos.Columns.Count - 1 do
            dTotalDia:=dTotalDia + StrToFloatDef(NxGridDatos.Cell[i,TRecurso(ListaRecursos.Objects[sDatosPanel.ListaRecursos.IndexOf(QrConcepto.FieldByName(ArrayTipo[iTipo,3]).AsString)]).iReng].AsString,0);

          if not NxGridTotales.RowExist(TRecurso(ListaRecursos.Objects[sDatosPanel.ListaRecursos.IndexOf(QrConcepto.FieldByName(ArrayTipo[iTipo,3]).AsString)]).iReng) then
            NxGridTotales.AddRow();

          NxGridTotales.Cell[0,TRecurso(ListaRecursos.Objects[sDatosPanel.ListaRecursos.IndexOf(QrConcepto.FieldByName(ArrayTipo[iTipo,3]).AsString)]).iReng].Asfloat:=dTotalDia;
          QrConcepto.Next;
        end;
      end;
    end;
    sDatosPanel.NxGridDatos.EndUpdate;

    if sDatosPanel.NxGridDatos.RowCount>24 then
    begin
      sDatosPanel.NxGridDatos.VertScrollBar.AutoHide:=False;
      sDatosPanel.NxGridDatos.VertScrollBar.Visible:=False;
      sDatosPanel.NxGridDatos.HorzScrollBar.AutoHide:=False;
      sDatosPanel.NxGridDatos.HorzScrollBar.Visible:=False;
      sDatosPanel. NxGridDatos.MouseWheelEnabled:=true;
    end
    else
      sDatosPanel.NxGridDatos.MouseWheelEnabled:=False;
  end
  else
    messageDLG('No se encontró personal el día anterior', mtInformation, [mbOk], 0);
end;

procedure TFrmCuadreXPartida.EliminarCategoria1Click(Sender: TObject);
var
  sDatosPanel:TSuperPanel;
  iPosDato:Integer;
  Recurso:TRecurso;
begin
  if MessageDlg('¿Desea Eliminar la Categoria?',mtConfirmation,[mbYes,mbNo],0)=mrYes then
  begin
    sDatosPanel:=TSuperPanel(ListaObj.Objects[ListaObj.IndexOf(NxPgcDatos.Pages[NxPgcDatos.ActivePageIndex].Name)]);
    iPosDato:=sDatosPanel.ListaRecursos.IndexOf(sDatosPanel.NxGridDatos.Cell[0,sDatosPanel.NxGridDatos.SelectedRow].AsString);
    if iPosDato<>-1 then
    begin
      Recurso:=TRecurso(sDatosPanel.ListaRecursos.Objects[iPosDato]);
      FreeAndNil(Recurso);
      sDatosPanel.ListaRecursos.Delete(iPosDato);
    end;
    sDatosPanel.NxGridDatos.DeleteRow(sDatosPanel.NxGridDatos.SelectedRow);

  end;
end;

procedure TFrmCuadreXPartida.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  OldFrente:='';
  OldCategoria:='';
  OldFecha:=0;
end;

procedure TFrmCuadreXPartida.FormCreate(Sender: TObject);
begin
  ListaObj:=TStringList.Create;
end;

procedure TFrmCuadreXPartida.FormDestroy(Sender: TObject);
var
  i:Integer;
begin
  for I := ListaObj.Count - 1 downto 0  do
    TSuperPanel(ListaObj.Objects[i]).Destroy;

  ListaObj.Destroy;
  inherited;
end;

procedure TFrmCuadreXPartida.FormShow(Sender: TObject);
begin
    //sIdPernocta:='CAESAR';
  isChange:=False;
  Cerrando:=False;
  If global_orden_general <> '' then
  Begin
      OrdenesdeTrabajo.Active := False ;
      OrdenesdeTrabajo.SQL.Clear ;
      OrdenesdeTrabajo.SQL.Add('Select ot.sNumeroOrden, ot.iJornada,Ot.sIdPernocta,Ot.sIdPlataforma from ordenesdetrabajo ot ' +
                               'INNER JOIN ordenesxusuario ou On (ou.sContrato=ot.sContrato '  +
                               'And ou.sNumeroOrden=ot.sNumeroOrden) ' +
                               'where ot.sContrato =:Contrato And ou.sDerechos<>"BLOQUEADO" ' +
                               'And sNumeroOrden = :orden ' +
                               'And ou.sIdUsuario =:Usuario And ot.cIdStatus =:Status order by ot.sNumeroOrden') ;
      OrdenesdeTrabajo.Params.ParamByName('Contrato').DataType := ftString ;
      OrdenesdeTrabajo.Params.ParamByName('Contrato').Value    := Global_Contrato ;
      OrdenesdeTrabajo.Params.ParamByName('orden').DataType    := ftString ;
      OrdenesdeTrabajo.Params.ParamByName('orden').Value       := global_orden_general ;
      OrdenesdeTrabajo.Params.ParamByName('Usuario').DataType  := ftString ;
      OrdenesdeTrabajo.Params.ParamByName('Usuario').Value     := Global_Usuario ;
      OrdenesdeTrabajo.Params.ParamByName('status').DataType   := ftString ;
      OrdenesdeTrabajo.Params.ParamByName('status').Value      := connection.configuracion.FieldValues [ 'cStatusProceso' ];

      OrdenesdeTrabajo.Open ;
  End
  Else
  Begin
      OrdenesdeTrabajo.Active := False ;
      OrdenesdeTrabajo.SQL.Clear ;
      if (global_grupo = 'INTEL-CODE') Then
          OrdenesdeTrabajo.SQL.Add('Select sNumeroOrden, iJornada,sIdPernocta,sIdPlataforma from ordenesdetrabajo where sContrato = :Contrato and ' +
                               'cIdStatus = :status order by sNumeroOrden')
      Else
      OrdenesdeTrabajo.SQL.Add('Select ot.sNumeroOrden, ot.iJornada,sIdPernocta,Ot.sIdPlataforma from ordenesdetrabajo ot ' +
                               'INNER JOIN ordenesxusuario ou On (ou.sContrato=ot.sContrato '  +
                               'And ou.sNumeroOrden=ot.sNumeroOrden) ' +
                               'where ot.sContrato =:Contrato And ou.sDerechos<>"BLOQUEADO" ' +
                               'And ou.sIdUsuario =:Usuario And ot.cIdStatus =:Status order by ot.sNumeroOrden') ;
      OrdenesdeTrabajo.Params.ParamByName('Contrato').DataType := ftString ;
      OrdenesdeTrabajo.Params.ParamByName('Contrato').Value    := Global_Contrato ;
      OrdenesdeTrabajo.Params.ParamByName('status').DataType   := ftString ;
      OrdenesdeTrabajo.Params.ParamByName('status').Value      := connection.configuracion.FieldValues [ 'cStatusProceso' ];
      if global_grupo <> 'INTEL-CODE' Then
        Begin
          OrdenesdeTrabajo.Params.ParamByName('Usuario').DataType  := ftString ;
          OrdenesdeTrabajo.Params.ParamByName('Usuario').Value     := Global_Usuario ;
        end;
      OrdenesdeTrabajo.Open ;
  End ;

  if sFrente<>'' then
   // tsNumeroOrden.KeyValue:=sFrente
  else
    if ordenesdetrabajo.RecordCount>0 then
   //   tsNumeroOrden.KeyValue:=OrdenesdeTrabajo.FieldByName('sNumeroOrden').AsString;

  if dFecha=0 then
    dFecha:=Now;

  tdIdFecha.Date:=dFecha;

  OldCategoria:='PERSONAL DE CONSTRUCCION';
 // OldFrente:=tsNumeroOrden.KeyValue;
  tsIdCategorias.Items.Clear;
  tsIdCategorias.Items.Add('PERSONAL DE CONSTRUCCION');
  tsIdCategorias.Items.Add('EQUIPO DE CONSTRUCCION');
  tsIdCategorias.Items.Add('PERSONAL DE ADMINISTRACION');
  tsIdCategorias.ItemIndex := 0;

  tdIdFechaExit(Sender);
end;

procedure TFrmCuadreXPartida.tdIdFechaExit(Sender: TObject);
begin
  tdidfecha.Color := global_color_salida;
  InicarCategoria;
  if OldFecha<>tdIdFecha.Date then
  begin
    OldFecha:=tdIdFecha.Date;
    tmrRecargar.Enabled:=True;
  end;
end;

procedure TFrmCuadreXPartida.tmrRecargarTimer(Sender: TObject);
begin
  tmrRecargar.Enabled:=False;
 // tsNumeroOrden.SetFocus;
  RecarGarDatos;
end;

procedure TFrmCuadreXPartida.tsIdCategoriasExit(Sender: TObject);
begin
  InicarCategoria;
  if OldCategoria<>tsIdCategorias.Text then
  begin
    OldCategoria:=tsIdCategorias.Text;
    tmrRecargar.Enabled:=True;
  end;
end;

procedure TFrmCuadreXPartida.tsNumeroOrdenExit(Sender: TObject);
begin
 { InicarCategoria;
  //if OldFrente<>tsNumeroOrden.KeyValue then
  begin
  //  OldFrente:=tsNumeroOrden.KeyValue;
    tmrRecargar.Enabled:=True;
  end; }
end;

end.
