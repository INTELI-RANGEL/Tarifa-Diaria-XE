unit Frm_ResumenPersonal;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, NxColumnClasses, NxColumns, NxScrollControl,
  NxCustomGridControl, NxCustomGrid, NxGrid, ComCtrls, ExtCtrls,global, DB,
  ZAbstractRODataset, ZAbstractDataset, ZDataset, Menus, AdvGlowButton, utilerias,
  NxEdit,nxclasses,StrUtils, AdvCombo;

type
  TTabla = (BPersonal,BPersonalCuadre);

  TModoImportacion = (Sustituir,Sumar,Agregar);

  TCategoria = class
    private
    Identificador,
    Descripcion: String;
  end;

  TtipoPernocta = class
    private
    Identificador,
    Descripcion: String;
  end;

  TFolio = class
    private
    Identificador,
    Descripcion,
    Plataforma,
    Pernoctan: string;
  end;

  TPernocta = class
    private
    Identificador,
    Descripcion: string;
  end;

  TPlataforma = class
    private
    Identificador,
    Descripcion: string;
  end;
  TFrmResumenPersonal = class(TForm)
    Panel1: TPanel;
    Panel2: TPanel;
    GroupBox1: TGroupBox;
    dIdFecha: TDateTimePicker;
    ComboBox: TComboBox;
    cmbCategoria: TComboBox;
    Label2: TLabel;
    Label1: TLabel;
    Label3: TLabel;
    NextGrid: TNextGrid;
    Procesar: TNxCheckBoxColumn;
    sidPersonal: TNxTextColumn;
    sTipoObra: TNxTextColumn;
    dCantidad: TNxNumberColumn;
    sDescripcion: TNxTextColumn;
    pBarraAvance: TProgressBar;
    PopupMenu1: TPopupMenu;
    Importardesdecuadre1: TMenuItem;
    Importardesdedaanterior1: TMenuItem;
    Sustituyendo1: TMenuItem;
    Sumando1: TMenuItem;
    Agregando1: TMenuItem;
    LTituloGrid: TLabel;
    btnPost: TAdvGlowButton;
    DesSeleccionartodo1: TMenuItem;
    N1: TMenuItem;
    dCantHH: TNxNumberColumn;
    dSolicitado: TNxNumberColumn;
    sHoraInicio: TNxTextColumn;
    ShoraFinal: TNxTextColumn;
    EditAjustaPernocta: TNxNumberEdit;
    Label4: TLabel;
    ClsTipoPernocta: TNxComboBoxColumn;
    sAgrupaPersonal: TNxTextColumn;
    ClFolio: TNxComboBoxColumn;
    ClsIdPernocta: TNxComboBoxColumn;
    ClsIdPlataforma: TNxComboBoxColumn;
    Imppersonaldediaanteriorcondiferentefolio1: TMenuItem;
    PnlDiaAnterior: TPanel;
    Panel4: TPanel;
    Panel5: TPanel;
    Panel6: TPanel;
    GrdPersonal: TNextGrid;
    AdvGlowButton1: TAdvGlowButton;
    CbxProcesar: TNxCheckBoxColumn;
    NxTextColumn1: TNxTextColumn;
    NxMemoColumn1: TNxMemoColumn;
    NxTextColumn2: TNxTextColumn;
    NxTreeColumn1: TNxTreeColumn;
    cmbChanFolio: TAdvComboBox;
    NxTextColumn3: TNxTextColumn;
    NxTextColumn4: TNxTextColumn;
    NxTextColumn5: TNxTextColumn;
    LbPlataforma: TLabel;
    AdvGlowButton2: TAdvGlowButton;
    LbPErnocta: TLabel;
    Sustituircambio: TCheckBox;
    Label5: TLabel;
    Label6: TLabel;
    ChbHD: TCheckBox;
    NxComboBoxColumn1: TNxComboBoxColumn;
    mniCambio: TMenuItem;
    procedure FormShow(Sender: TObject);
    procedure cmbCategoriaChange(Sender: TObject);
    procedure dIdFechaChange(Sender: TObject);
    procedure ComboBoxChange(Sender: TObject);
    procedure NextGridAfterEdit(Sender: TObject; ACol, ARow: Integer;
      Value: WideString);
    procedure Sustituyendo1Click(Sender: TObject);
    procedure Sumando1Click(Sender: TObject);
    procedure Agregando1Click(Sender: TObject);
    procedure Importardesdedaanterior1Click(Sender: TObject);
    procedure btnPostClick(Sender: TObject);
    procedure DesSeleccionartodo1Click(Sender: TObject);
    procedure Imppersonaldediaanteriorcondiferentefolio1Click(Sender: TObject);
    procedure CbxProcesarChange(Sender: TObject);
    procedure AdvGlowButton2Click(Sender: TObject);
    procedure cmbChanFolioChange(Sender: TObject);
    procedure AdvGlowButton1Click(Sender: TObject);
    procedure ChbHDClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure PopupMenu1Popup(Sender: TObject);
  private
    procedure CargarCategorias(Cmb: Tcombobox);
    procedure CargarPersonalGrid(Categoria: string; Fecha: Tdatetime;Grid:TnextGrid);
    procedure CargarFoliosXFecha(Fecha: Tdatetime);
    procedure CargarCantidades(Grid: TnextGrid; Fecha: Tdatetime; Categoria,
      Orden: String);
    procedure GuardarCambios(Grid:TnextGrid;Categoria,Folio:String;Fecha:TdateTime);
    procedure ImportarPersonal(Desde: TTabla; Modo: TModoImportacion;Fecha:Tdatetime);
    procedure DesSeleccionarTodo(Grid: TnextGrid;TP:Boolean = False);
    procedure mniMiCambioClick(Sender: TObject);

    
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FrmResumenPersonal: TFrmResumenPersonal;


const
  {$Region 'Constantes de Columnas por campo (Referencias)'}
  CProcesar = 0;
  CsIdPersonal = 1;
  CsTipoObra = 2;
  CsDescripcion = 3;
  CsIdPernocta = 4;
  CsIdPlataforma = 5;
  CsNumeroOrden = 6;
  CdCantidad = 7;
  CsAgrupaPersonal = 8;
  CdCantHH = 9;
  CdSolicitado = 10;
  CsHoraInicio = 11;
  CsHoraFinal = 12;
  CsTipoPernocta = 13;
  {$EndRegion}

implementation
uses frm_connection, UnitTarifa, UFunctionsGHH;

{$R *.dfm}

{$REGION 'INTERFAZ'}
procedure TFrmResumenPersonal.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  Action := caFree;
end;
procedure TFrmResumenPersonal.mniMiCambioClick(Sender: TObject);
var
  SFolio,sFolioOrigen:String;
  QSource,QTarget:TzQuery;
begin
  if MessageDlg('Se van a Guardar los cambios al Sistema.' + #13 + #10 +
  '¿Desea Continuar?',mtConfirmation,[MbYes,MbNo],0 )=MrYes then
  begin
    GuardarCambios(NextGrid,TCategoria(cmbcategoria.Items.Objects[cmbCategoria.ItemIndex]).Identificador,Combobox.text,didfecha.datetime);

    if (sender as TMenuItem).Caption='&NA' then
      sFolio:='@'
    else
      sFolio:=AnsiReplaceText((sender as TMenuItem).Caption,'&','');

    //AnsiReplaceText(aSql[iCiclo,1],'group by',aSql[iCiclo,2]);
    if Combobox.Text='NA' then
      sFolioOrigen:='@'
    else
      sFolioOrigen:=Combobox.Text;

    QSource:=TzQuery.Create(nil);
    QTarget:=TzQuery.Create(nil);
    try
      try
        QSource.Connection:=Connection.zConnection;
        QSource.SQL.Text:='select bpc.* from bitacoradepersonal_cuadre bpc '+
                          'inner join personal p on ( p.sContrato =:ContratoBarco and bpc.sIdPersonal = p.sIdPersonal ) '+
                          'inner join tiposdepersonal tp on (tp.sIdTipoPersonal = p.sIdTipoPersonal ) '+
                          'where bpc.sContrato=:Contrato and ' +
                          'bpc.dIdFecha=:fecha and bpc.sNumeroOrden=:Folio and tp.sIdTipoPersonal=:Tipo';
        QSource.ParamByName('Contrato').AsString:=global_Contrato;
        QSource.ParamByName('ContratoBarco').AsString:=global_Contrato_Barco;
        QSource.ParamByName('Fecha').AsDate:=dIdFecha.Date;
        QSource.ParamByName('Folio').AsString:=sFolioOrigen;
        QSource.ParamByName('Tipo').AsString:=TCategoria(CmbCategoria.Items.Objects[CmbCategoria.ItemIndex]).Identificador;
        QSource.Open;

        QTarget.Connection:=Connection.zConnection;
        QTarget.SQL.Text:='select bpc.* from bitacoradepersonal_cuadre bpc '  +
                          'inner join personal p on ( p.sContrato =:ContratoBarco and bpc.sIdPersonal = p.sIdPersonal ) '+
                          'inner join tiposdepersonal tp on (tp.sIdTipoPersonal = p.sIdTipoPersonal ) '+
                          'where bpc.sContrato=:Contrato and ' +
                          'bpc.dIdFecha=:fecha and bpc.sNumeroOrden=:Folio and tp.sIdTipoPersonal=:Tipo';
        QTarget.ParamByName('Contrato').AsString:=global_Contrato;
        QTarget.ParamByName('ContratoBarco').AsString:=global_Contrato_Barco;
        QTarget.ParamByName('Fecha').AsDate:=dIdFecha.Date;
        QTarget.ParamByName('Folio').AsString:=sFolio;
        QTarget.ParamByName('Tipo').AsString:=TCategoria(CmbCategoria.Items.Objects[CmbCategoria.ItemIndex]).Identificador;
        QTarget.Open;

        while not QSource.Eof do
        begin
          if QTarget.Locate('sIdPersonal',QSource.FieldByName('sIdPersonal').AsString,[]) then
          begin
            QTarget.Edit;
            QTarget.FieldByName('dCantidad').AsFloat:= QTarget.FieldByName('dCantidad').AsFloat + QSource.FieldByName('dCantidad').AsFloat;
            QTarget.Post;
          end
          else
          begin
            QTarget.Append;
            MyCopyFields(QTarget,QSource);
            QTarget.FieldByName('sNumeroOrden').AsString:=sFolio;
            QTarget.Post;
          end;
          QSource.Next;
        end;

        QSource.First;
        while not QSource.Eof do
          QSource.Delete;

        TdProrrateaPernocta(global_Contrato,didfecha.Date);
        CargarCantidades(NextGrid,didfecha.datetime,Tcategoria(cmbCategoria.Items.Objects[cmbCategoria.ItemIndex]).Identificador,combobox.text);
      except
        on e:Exception do
          MessageDlg(e.Message + ', ' + e.Message,MtInformation,[Mbok],0 );
      end;
    finally
      QSource.Destroy;
      QTarget.Destroy;
    end;
  end
  else
    MessageDlg('Operacion Cancelada.',MtInformation,[Mbok],0 );
end;

procedure TFrmResumenPersonal.FormShow(Sender: TObject);
begin
  self.Caption := 'R E S U M E N   D E   P E R S O N A L   P A R A   E L   C O N T R A T O : '+Global_Contrato;
  dIdFecha.DateTime := now-1;
  CargarCategorias(cmbCategoria);
  CargarPersonalGrid(TCategoria(CmbCategoria.Items.Objects[CmbCategoria.ItemIndex]).Identificador,DidFecha.DateTime,NextGrid);
  CargarFoliosXFecha(DidFecha.DateTime);
  CargarCantidades(NextGrid,didfecha.datetime,Tcategoria(cmbCategoria.Items.Objects[cmbCategoria.ItemIndex]).Identificador,combobox.text);
end;
procedure TFrmResumenPersonal.cmbCategoriaChange(Sender: TObject);
begin

  CargarPersonalGrid(TCategoria(CmbCategoria.Items.Objects[CmbCategoria.ItemIndex]).Identificador,DidFecha.DateTime,NextGrid);
  CargarFoliosXFecha(DidFecha.DateTime);
  CargarCantidades(NextGrid,didfecha.datetime,Tcategoria(cmbCategoria.Items.Objects[cmbCategoria.ItemIndex]).Identificador,combobox.text);
end;
procedure TFrmResumenPersonal.cmbChanFolioChange(Sender: TObject);
begin
  try
    LbPlataforma.Caption := tfolio(cmbChanFolio.Items.Objects[cmbChanFolio.ItemIndex]).Plataforma;
    LbPErnocta.Caption :=tfolio(cmbChanFolio.Items.Objects[cmbChanFolio.ItemIndex]).Pernoctan;
  except
    ;
  end;
end;

procedure TFrmResumenPersonal.ComboBoxChange(Sender: TObject);
begin
  CargarCantidades(NextGrid,didfecha.datetime,Tcategoria(cmbCategoria.Items.Objects[cmbCategoria.ItemIndex]).Identificador,combobox.text);
end;

procedure TFrmResumenPersonal.dIdFechaChange(Sender: TObject);
begin
  CargarFoliosXFecha(DidFecha.DateTime);
  CargarCantidades(NextGrid,didfecha.datetime,Tcategoria(cmbCategoria.Items.Objects[cmbCategoria.ItemIndex]).Identificador,combobox.text);
end;

procedure TFrmResumenPersonal.NextGridAfterEdit(Sender: TObject; ACol,
  ARow: Integer; Value: WideString);
begin
  if ACol = CdCantidad then
  begin
    if NextGrid.Cell[ACol, ARow].AsFloat > 0 then
    begin
      if NextGrid.Cell[CProcesar, ARow].AsBoolean=false then
        NextGrid.Cell[14, ARow].AsString:='Si';

      NextGrid.Cell[CProcesar, ARow].AsBoolean := True;

    end
    else
    begin
      if length(trim(NextGrid.Cells[CsTipoObra, ARow])) = 0 then
      begin
        NextGrid.Cell[CProcesar, ARow].AsBoolean := False;
        NextGrid.Cell[14, ARow].AsString:='No';
      end;
    end;
  end;
  NextGrid.CalculateFooter();

end;

procedure TFrmResumenPersonal.PopupMenu1Popup(Sender: TObject);
var
  i:Integer;
begin
  mniCambio.Enabled:=false;
  if combobox.Items.Count>1 then
  begin
    mniCambio.Enabled:=true;
    for I := 0 to mniCambio.Count - 1 do
      mniCambio.Items[i].Enabled:=true;

    mniCambio.Find(combobox.Text).Enabled:=false;
  end;

end;

procedure TFrmResumenPersonal.DesSeleccionartodo1Click(Sender: TObject);
begin
  DesSeleccionarTodo(NextGrid);
end;

procedure TFrmResumenPersonal.Sumando1Click(Sender: TObject);
begin
  ImportarPersonal(BPersonal,Sumar,didfecha.Datetime);
  CargarCantidades(NextGrid,didfecha.datetime,Tcategoria(cmbCategoria.Items.Objects[cmbCategoria.ItemIndex]).Identificador,combobox.text);
end;

procedure TFrmResumenPersonal.Sustituyendo1Click(Sender: TObject);
begin
  ImportarPersonal(BPersonal,Sustituir,didfecha.Datetime);
  CargarCantidades(NextGrid,didfecha.datetime,Tcategoria(cmbCategoria.Items.Objects[cmbCategoria.ItemIndex]).Identificador,combobox.text);
end;

procedure TFrmResumenPersonal.AdvGlowButton1Click(Sender: TObject);
var
  updbpc,delbpc:TZQuery;
  i:Integer;
begin
  updbpc := TZQuery.Create(nil);
  try
    updbpc.Connection := connection.zConnection;
    updbpc.Active := False;
    updbpc.SQL.Clear;
    updbpc.SQL.Text :=
                       'INSERT IGNORE INTO bitacoradepersonal_cuadre ' +
                       ' ( sContrato,  sNumeroOrden,  dIdFecha,  sIdPersonal,sTipoObra,  sDescripcion,  sIdPernocta,  sIdPlataforma,  sHoraInicio,  sHoraFinal,  dCantidad, iitemorden,  sagrupapersonal,  laplicapernocta, dCantHH, dSolicitado,sTipoPernocta)' +
                       ' VALUES ' +
                       ' (:sContrato, :sNumeroOrden, :dIdFecha, :sIdPersonal,:sTipoObra, :sDescripcion, :sIdPernocta, :sIdPlataforma, :sHoraInicio, :sHoraFinal, :dCantidad,:iitemorden, :sagrupapersonal, :laplicapernocta, :dCantHH, :dSolicitado, :sTipoPernocta); ' ;

    if GrdPersonal.RowCount = 0 then
      raise exception.Create('No hay personal cargado');




    try
      updbpc.Connection.StartTransaction;

      if Sustituircambio.Checked then
      begin
        delbpc := TZQuery.Create(nil);
        try
          delbpc.Connection := connection.zConnection;
          delbpc.Active := False;
          delbpc.SQL.Clear;
          delbpc.SQL.Text := 'delete from bitacoradepersonal_cuadre where scontrato = :Contrato and didfecha = :fecha and snumeroorden = :folio ';
          delbpc.ParamByName('contrato').AsString := global_contrato;
          delbpc.ParamByName('folio').AsString := cmbChanFolio.Text;
          delbpc.ParamByName('fecha').AsDate := dIdFecha.Date;
          delbpc.ExecSQL;
        finally
          delbpc.Free;
        end;
      end;

      for I := 0 to GrdPersonal.RowCount-1 do
      begin
        if (GrdPersonal.Row[i].Level = 1) and (GrdPersonal.Cell[2,i].AsBoolean = True) then
        begin
          updbpc.ParamByName('scontrato').AsString := global_contrato;
          updbpc.ParamByName('snumeroorden').AsString := cmbChanFolio.Text;
          updbpc.ParamByName('didfecha').AsDate := dIdFecha.Date;
          updbpc.ParamByName('sidpersonal').AsString := GrdPersonal.Cell[4,i].AsString;
          updbpc.ParamByName('stipoobra').AsString :=GrdPersonal.Cell[5,i].AsString;
          updbpc.ParamByName('sdescripcion').AsString :=GrdPersonal.Cell[6,i].AsString;
          updbpc.ParamByName('sidpernocta').AsString := LbPErnocta.Caption;
          updbpc.ParamByName('sidplataforma').AsString := LbPlataforma.Caption;
          updbpc.ParamByName('shorainicio').AsString := '00:00';
          updbpc.ParamByName('shorafinal').AsString := '00:00';
          updbpc.ParamByName('dcantidad').AsFloat := GrdPersonal.Cell[3,i].AsFloat;
          updbpc.ParamByName('iitemorden').AsString := '0';
          updbpc.ParamByName('sagrupapersonal').AsString :='@';
          updbpc.ParamByName('laplicapernocta').AsString :='Si';
          updbpc.ParamByName('dCantHH').AsFloat :=0;
          updbpc.ParamByName('dSolicitado').AsFloat :=0;
          updbpc.ParamByName('sTipoPernocta').AsString :=GrdPersonal.Cell[7,i].AsString;
          updbpc.ExecSQL;
        end;
      end;
      if updbpc.Connection.InTransaction then
        updbpc.Connection.Commit;
      PnlDiaAnterior.Visible := False;  
    except
      on e:Exception do
      begin
        if updbpc.Connection.InTransaction then
          updbpc.Connection.Rollback;
        ShowMessage('Ocurrio el siguiente error al tratar de guardar la informacvión: '+#10+e.Message);
      end;
    end;

  finally
    updbpc.Free;
  end;
end;

procedure TFrmResumenPersonal.AdvGlowButton2Click(Sender: TObject);
begin
  PnlDiaAnterior.Visible := False;
end;

procedure TFrmResumenPersonal.Agregando1Click(Sender: TObject);
begin
  ImportarPersonal(BPersonal,Agregar,didfecha.Datetime);
  CargarCantidades(NextGrid,didfecha.datetime,Tcategoria(cmbCategoria.Items.Objects[cmbCategoria.ItemIndex]).Identificador,combobox.text);
end;

procedure TFrmResumenPersonal.btnPostClick(Sender: TObject);
begin
  btnpost.setfocus;
  GuardarCambios(NextGrid,TCategoria(cmbcategoria.Items.Objects[cmbCategoria.ItemIndex]).Identificador,Combobox.text,didfecha.datetime);
  TdProrrateaPernocta(global_Contrato,didfecha.Date);

end;

procedure TFrmResumenPersonal.Importardesdedaanterior1Click(Sender: TObject);
begin
  ImportarPersonal(BPersonalCuadre,Sustituir,didfecha.Datetime);
  CargarCantidades(NextGrid,didfecha.datetime,Tcategoria(cmbCategoria.Items.Objects[cmbCategoria.ItemIndex]).Identificador,combobox.text);
end;
{$ENDREGION}

{$REGION 'PROCEDIMIENTOS Y FUNCIONES'}
procedure TFrmResumenPersonal.CargarCategorias(Cmb:Tcombobox);
{
  CARGAR LAS CATEGORIAS DESDE LA TABLA TIPOS DE PERSONAL A UN COMBOBOX
}
Var
  Query: TZQuery;
  CrLCat:Tcursor;
  Categoria:Tcategoria;
begin
  try
    Query := TZQuery.Create(nil);
    CrLCat := screen.Cursor;
    screen.Cursor := crappstart;
    try
      Query.Connection := Connection.zConnection;
      Query.SQL.Text := 'SELECT * FROM tiposdepersonal;';   
      Query.Open;
      Cmb.Items.Clear;
      while Not Query.Eof do
      begin
        Categoria := Tcategoria.create;
        Categoria.Identificador := Query.FieldByName('sIdTipoPersonal').AsString;
        Categoria.Descripcion := Query.FieldByName('sDescripcion').AsString;
        Cmb.Items.AddObject(Categoria.Identificador + ' | ' + Categoria.Descripcion,Categoria);
        Query.Next;
      end;
      if cmb.Items.count > 0 then
        Cmb.ItemIndex := 0;
    finally
      screen.Cursor := CrLCat;
      Query.Free;
    end;
  except
    on e:Exception do
      showmessage('Ocurrió el siguiente error al cargar las categorias del personal: '+#10+e.message);
  end;
end;

procedure TFrmResumenPersonal.CargarPersonalGrid(Categoria:string;Fecha:Tdatetime;Grid:TnextGrid);
{
  CARGAR LOS PERSONALES DADA UNA CATEGORIA DE PERSONAL, ESTO DEPENDE DEL TIPO DE
  CATEGORIA QUE SE SELECCIONE, EN CASO DE QUE SEA PU SE CARGARA DESDE MOERECURSOS
  TOMANDO EN CUENTA EL MOE MAS PROXIMO A LA FECHA.
  EN CASO DE QUE SEA DIFERENTE DE PU SE DEBE MOSTRAR EL PERSONAL DESDE EL CATALOGO
  DE PERSONAL
  POSTERIORMENTE SE CARGA EL GRID
}
var
  QryPersonal: TZQuery;
  CurPersonal:Tcursor;
  TipoPErnocta: TtipoPernocta;
  Pernocta: TPernocta;
  Plataforma:TPlataforma;
begin
  //si es personal de la orden entonces no debes elegir cual es el folio por personal
  //pero queda visible para que se puedan asignar los ajustes por percnotas a folio
  //aplicables solo al personalde pu
  if Categoria <> '.' then
  begin
    ClFolio.visible := False;
    editajustapernocta.visible := True;
    label4.Visible := True;
  end
  else
  begin
    ClFolio.visible := True;
    editajustapernocta.visible := False;
    label4.Visible := False;
  end;
  try
    QryPersonal := TZQuery.Create(Self);
    CurPersonal := screen.Cursor;
    screen.cursor := crappstart;
    try
      QryPersonal.Connection := Connection.zConnection;
      if Categoria = 'PU' then
      begin
        QryPersonal.Active := False;
        QryPersonal.SQL.Clear;
        QryPersonal.SQL.Add('' +
                      'SELECT ' +
                      ' p.sDescripcion AS sCategoria, ' +
                      ' p.sIdPersonal, ' +
                      ' p.sDescripcion ' +
                      'FROM moe AS m ' +
                      '	INNER JOIN moerecursos AS mr ' +
                      '		ON (mr.iidMoe=m.iidMoe) ' +
                      '	INNER JOIN personal AS p ' +
                      '		ON (p.scontrato=:Contrato AND p.sidpersonal=mr.sidRecurso) ' +
                      'left JOIN gruposdepersonal AS gp '+
                      '	ON (gp.iId = p.iId_AgrupadorPersonal) '+
                      'WHERE ' +
                      '	m.didfecha = (SELECT max(didfecha) FROM moe WHERE didfecha <= :Fecha AND sContrato = :ContratoNormal) ' +
                      ' AND m.sContrato = :ContratoNormal ' +
                      '	AND mr.eTipoRecurso = "Personal" ' +
                      'GROUP BY p.sidpersonal '+
                      ' ORDER BY p.iItemOrden ASC');
                     { 'SELECT  '+
                      'p.iId_AgrupadorPersonal,  '+
                      'p.sDescripcion AS sCategoria, '+
                      'p.sIdPersonal,   '+
                      'p.sDescripcion,  '+
                      'if(SUM(bp.dCantHH) > 0, SUM(bp.dCantHH), SUM(bp.dCantidad)) AS dCantidadABordo, '+
                      '( SELECT       '+
                      'SUM(mr.dCantidad)  '+
                      'FROM moerecursos AS mr  '+
                      '	INNER JOIN personal AS per   '+
                      '	ON (per.sIdPersonal = mr.sIdRecurso AND per.sContrato = :Contrato) '+
                      '	WHERE   '+
                      '	mr.eTipoRecurso = "Personal"  '+
                      '	AND mr.iIdMoe = m.iIdMoe '+
                      '		AND per.iId_AgrupadorPersonal = p.iId_AgrupadorPersonal '+
                      '   AND per.lSumaSolicitado = "Si"   '+
                      '	) AS dCantidadSolicitada '+
                      'FROM bitacoradepersonal_cuadre AS bp  '+
                    	'INNER JOIN personal AS p '+
                      '	ON (p.sIdPersonal = bp.sIdPersonal AND p.sContrato = :Contrato) '+
                      'INNER JOIN gruposdepersonal AS gp '+
                      '	ON (gp.iId = p.iId_AgrupadorPersonal) '+
                      '	INNER JOIN moe AS m '+
                      '		ON (m.dIdFecha = (SELECT max(dIdFecha) FROM moe WHERE dIdFecha <= :Fecha AND sContrato = :ContratoNormal) AND m.sContrato = :ContratoNormal) '+
                      'WHERE bp.dIdFecha = :Fecha '+
                      'AND bp.sContrato = :ContratoNormal '+
                      'GROUP BY gp.sGrupo '+
                      'ORDER BY p.iItemOrden ASC; ');   }
        QryPersonal.ParamByName('Contrato').AsString := Global_Contrato_Barco;
        QryPersonal.ParamByName('ContratoNormal').AsString := Global_Contrato;
        QryPersonal.ParamByName('Fecha').AsString := FormatDateTime('yyyy-mm-dd', Fecha);
      end
      else
      begin
        QryPersonal.Active := False;
        QryPersonal.SQL.Clear;
        QryPersonal.SQL.Text := 'SELECT tp.sDescripcion AS sCategoria, p.sIdPersonal, p.sDescripcion FROM personal AS p ' +
                                'INNER JOIN tiposdepersonal AS tp ON (tp.sIdTipoPersonal = p.sIdTipoPersonal) ' +
                                'WHERE p.sContrato = :ContratoBarco AND p.sIdTipoPersonal = :IdPersonal ORDER BY p.sIdTipoPersonal, p.iItemOrden;';
        QryPersonal.Params.ParamByName('ContratoBarco').AsString := Global_Contrato_Barco;
        QryPersonal.Params.ParamByName('IdPersonal').AsString := Categoria;
      end;
      QryPersonal.open;

      //llenamos los combos para seleccionar el tipo de pernocta
      Connection.QryBusca.Active := False;
      Connection.QryBusca.SQL.Text := 'SELECT * FROM cuentas;';
      Connection.QryBusca.Open;
      ClStipoPernocta.DataMode := dmValueList;
      ClStipoPernocta.Items.Clear;
      while Not Connection.QryBusca.Eof do
      begin
        TipoPernocta := TtipoPernocta.Create;
        TipoPErnocta.Identificador := Connection.QryBusca.FieldByName('sIdCuenta').AsString;
        Tipopernocta.Descripcion := Connection.QryBusca.FieldByName('sDescripcion').AsString;
        ClStipoPernocta.Items.AddObject(' '+TipoPErnocta.Identificador+' | '+Tipopernocta.Descripcion,TipoPernocta);
        Connection.QryBusca.Next;
      end;

      //llenamos los combos para seleccionar la pernocta
      Connection.QryBusca.Active := False;
      Connection.QryBusca.SQL.Text := 'SELECT sIdPernocta, sDescripcion FROM pernoctan  Order By sDescripcion;';
      Connection.QryBusca.Open;
      ClsIdPernocta.DataMode := dmValueList;
      ClsIdPernocta.Items.Clear;
      while Not Connection.QryBusca.Eof do
      begin
        Pernocta := TPernocta.Create;
        Pernocta.Identificador := Connection.QryBusca.FieldByName('sIdPernocta').AsString;
        pernocta.Descripcion := Connection.QryBusca.FieldByName('sDescripcion').AsString;
        ClsIdPernocta.Items.AddObject(' '+pernocta.Descripcion,Pernocta);
        Connection.QryBusca.Next;
      end;

      //llenamos los combos para seleccionar la plataforma
      Connection.QryBusca.Active := False;
      Connection.QryBusca.SQL.Text := 'SELECT sIdPlataforma, sDescripcion FROM plataformas WHERE lStatus = "Activa" Order By sDescripcion;';
      Connection.QryBusca.Open;
      ClsIdPlataforma.DataMode := dmValueList;
      ClsIdPlataforma.Items.Clear;
      while Not Connection.QryBusca.Eof do
      begin
        Plataforma:= TPlataforma.Create;
        Plataforma.Identificador := Connection.QryBusca.FieldByName('sIdPlataforma').AsString;
        Plataforma.Descripcion := Connection.QryBusca.FieldByName('sDescripcion').AsString;
        ClsIdPlataforma.Items.AddObject(' '+Plataforma.Descripcion,Plataforma);
        Connection.QryBusca.Next;
      end;

      //llenamos el grid
      Grid.ClearRows;
      while Not QryPersonal.Eof do
      begin
        Grid.AddRow;
        Grid.Cells[CProcesar, Grid.LastAddedRow]    := 'False';
        Grid.Cells[CsIdPersonal, Grid.LastAddedRow] := QryPersonal.FieldByName('sIdPersonal').AsString;
        Grid.Cells[CsDescripcion, Grid.LastAddedRow]:= QryPersonal.FieldByName('sDescripcion').AsString;
        Grid.Cells[CdCantidad,Grid.LastAddedRow]    := '0';

        Grid.Cells[CsNumeroOrden,Grid.LastAddedRow]:= 'NA';

        if ClsTipoPErnocta.Items.Count > 0 then
          Grid.Cells[CsTipoPernocta,Grid.LastAddedRow]:= ClsTipoPErnocta.Items[0]
        else
          Grid.Cells[CsTipoPernocta,Grid.LastAddedRow]:= '';

        if ClsidPernocta.Items.Count > 0 then
          Grid.Cells[CsidPernocta,Grid.LastAddedRow]:= ClsidPernocta.Items[0]
        else
          Grid.Cells[CsidPernocta,Grid.LastAddedRow]:= '';

        if ClsidPlataforma.Items.Count > 0 then
          Grid.Cells[CsidPlataforma,Grid.LastAddedRow]:= ClsidPlataforma.Items[0]
        else
          Grid.Cells[CsidPlataforma,Grid.LastAddedRow]:= '';
        QryPersonal.Next;
      end;
    finally
      Screen.cursor := CurPersonal;
      QryPersonal.free;
    end;
  except
    on e:exception do
      showmessage('Ocurrió el siguiente error al tratar de cargar el personal al grid de datos: '+#10+e.message);
  end;
end;

procedure TFrmResumenPersonal.CbxProcesarChange(Sender: TObject);
var rw:Integer;

begin
  if GrdPersonal.Row[GrdPersonal.SelectedRow].Level = 0 then
  begin
    for rw := 0 to GrdPersonal.Row[GrdPersonal.SelectedRow].ChildCount do
      GrdPersonal.Cell[2,GrdPersonal.SelectedRow+rw].AsBoolean := GrdPersonal.Cell[2,GrdPersonal.SelectedRow].AsBoolean;
  end;
end;

procedure TFrmResumenPersonal.ChbHDClick(Sender: TObject);
begin
  DesSeleccionarTodo(NextGrid,ChbHd.Checked);
end;

procedure TFrmResumenPersonal.CargarFoliosXFecha(Fecha:Tdatetime);
{
  EN CASO DE QUE LA FECHA SELECCIONADA CONTENGA REGISTROS EN LA TABLA
  BITACORADEACTIVIDADES SE DEBE OBTENER LOS DIFERENTES FOLIOS Y CARGARLOS
  A UN COMBO, DE TAL FORMA QUE NOS PERMITA HACER UN FILTRADO
}
Var
  QueryX: TZQuery;
  CurFolios:Tcursor;
  MiItem:TMenuItem;
  i:Integer;
begin
  try
    CurFolios := Screen.Cursor;
    QueryX := TZQuery.Create(nil);
    try
      Screen.Cursor := Crappstart;
      QueryX.Connection := Connection.zConnection;

      QueryX.Active := False;
      QueryX.SQL.Text :=  '' +
                          'SELECT ' +
                          '	sNumeroOrden ' +
                          'FROM bitacoradeactividades AS ba ' +
                          'WHERE ' +
                          '	ba.sContrato = :Contrato ' +
                          '	AND ba.dIdFecha = :Fecha ' +
                          'GROUP BY sNumeroOrden;';
      QueryX.Params.ParamByName('Contrato').AsString := Global_Contrato;
      QueryX.Params.ParamByName('Fecha').AsString := FormatDateTime('yyyy-mm-dd', Fecha);
      QueryX.Open;
      QueryX.First;

      while mniCambio.Count>0 do
        mniCambio.Items[0].Destroy;
      ComboBox.Items.Clear;
      if tcategoria(cmbcategoria.Items.Objects[cmbcategoria.ItemIndex]).Identificador <> 'PU' then
      begin
        ComboBox.Items.add('NA');

        MiItem:=TMenuItem.Create(nil);
        MiItem.OnClick:=mniMiCambioClick;
        MiItem.Caption:='NA';
        mniCambio.Add(MiItem);
      end;

      while Not QueryX.Eof do
      begin
        ComboBox.Items.Add(QueryX.FieldByName('sNumeroOrden').AsString);
        MiItem:=TMenuItem.Create(nil);
        MiItem.OnClick:=mniMiCambioClick;
        MiItem.Caption:=QueryX.FieldByName('sNumeroOrden').AsString;
        mniCambio.Add(MiItem);

        QueryX.Next;
      end;
      ComboBox.ItemIndex := 0;
    finally
      Screen.Cursor := CurFolios;
      QueryX.free;
    end;
  except
    on e:Exception do
      showmessage('Ocurrió el siguiente error al tratar de cargar los folios: '+#10+e.message);
  end;
end;

procedure TFrmResumenPersonal.CargarCantidades(Grid:TnextGrid;Fecha:Tdatetime;Categoria,Orden:String);
{
  SE DEBE LEER DE LA BITACORADEPERSONALCUADRE TODOS LOS REGISTROS FILTRADOS POR FECHA, CATEGORIA Y FOLIO
  LUEGO SE DEBE IR RECORRIENDO EL GRID PARA ESTABLECER CANTIDADES LEIDAS DESDE ESA TABLA
}
Var
  Query: TZQuery;
  i:integer;
  CurCantidades:Tcursor;

Function LocalizaXId(IdentKey:String ):Integer;
var C,R:integer;
begin
  R := -1;
  for c := 0 to ClsTipoPernocta.items.Count -1 do
  begin
    if clsTipoPernocta.Items.Objects[c] <> nil then
      if TtipoPernocta(clsTipoPernocta.Items.objects[c]).Identificador = IdentKey then
        R := c;
  end;
  Result := R;
end;

Function LocalizaPernoctaXId(IdentKey:String ):Integer;
var C,R:integer;
begin
  R := -1;
  for c := 0 to ClsIdPernocta.items.Count -1 do
  begin
    if ClsIdPernocta.Items.Objects[c] <> nil then
      if TPernocta(ClsIdPernocta.Items.objects[c]).Identificador = IdentKey then
        R := c;
  end;
  Result := R;
end;

Function LocalizaPlataformaXId(IdentKey:String ):Integer;
var C,R:integer;
begin
  R := -1;
  for c := 0 to ClsIdPlataforma.items.Count -1 do
  begin
    if ClsIdPlataforma.Items.Objects[c] <> nil then
      if TPlataforma(ClsIdPlataforma.Items.objects[c]).Identificador = IdentKey then
        R := c;
  end;
  Result := R;
end;

begin
  try
    if Orden = 'NA' then
      Orden := '@';
    EditAjustaPernocta.Value := 0.000000;
    Query := TZQuery.Create(nil);
    CurCantidades := screen.Cursor;
    try
      screen.Cursor := CrAppStart;
      Query.Connection := Connection.zConnection;
      Query.SQL.Text := 'SELECT * FROM bitacoradepersonal_cuadre '+
                        'WHERE dIdFecha = :dIdFecha AND sTipoObra = :sTipoObra AND sContrato = :sContrato '+
                        'AND (:sNumeroOrden = "<TODOS>" or (:sNumeroOrden <> "<TODOS>" and  sNumeroOrden = :sNumeroOrden))';
      Query.Params.ParamByName('sContrato').AsString := Global_Contrato;
      Query.Params.ParamByName('dIdFecha').AsString := FormatDateTime('yyyy-mm-dd', Fecha);
      Query.Params.ParamByName('sTipoObra').AsString := Categoria;
      Query.Params.ParamByName('sNumeroOrden').AsString := Orden;
      Query.Open;

      ClFolio.Items.Clear;
     // if Categoria <> 'PU' then
      begin
        //llenamos los combos para seleccionar el folio
        Connection.QryBusca.Active := False;
        Connection.QryBusca.SQL.Text := ''+
                            'SELECT ' +
                            '	sNumeroOrden ' +
                            'FROM bitacoradeactividades AS ba ' +
                            'WHERE ' +
                            '	ba.sContrato = :Contrato ' +
                            '	AND ba.dIdFecha = :Fecha ' +
                            'GROUP BY sNumeroOrden;';
        Connection.QryBusca.Params.ParamByName('Contrato').AsString := Global_Contrato;
        Connection.QryBusca.Params.ParamByName('Fecha').AsString := FormatDateTime('yyyy-mm-dd', Fecha);
        Connection.QryBusca.Open;
        ClFolio.DataMode := dmValueList;
        ClFolio.Items.Add(' NA');
        while Not Connection.QryBusca.Eof do
        begin
          ClFolio.Items.Add(' '+Connection.QryBusca.FieldByName('sNumeroOrden').AsString);
          Connection.QryBusca.Next;
        end;
      end;


      for i := 0 to Grid.RowCount - 1 do
      begin
        if Query.Locate('sIdPersonal', Grid.Cells[1, i], []) then
        begin
          Grid.Cells[CProcesar, i] := 'True';
          Grid.Cells[CsTipoObra, i] := Query.FieldByName('stipoobra').asstring;
          //Grid.Cells[CsIdPernocta, i] := Query.FieldByName('sIdPernocta').asstring;
          if LocalizaPernoctaXId(Query.FieldByName('sIdPernocta').asstring) <> -1 then
            Grid.Cells[CsidPernocta, i] := ClsidPErnocta.Items[LocalizaPernoctaXId(Query.FieldByName('sIdPernocta').asstring)]
          else
            Grid.Cells[CsidPernocta, i] := '';
         // Grid.Cells[CsIdPlataforma, i] := Query.FieldByName('sIdPlataforma').asstring;
          if LocalizaPlataformaXId(Query.FieldByName('sIdPlataforma').asstring) <> -1 then
            Grid.Cells[CsidPlataforma, i] := ClsidPlataforma.Items[LocalizaPlataformaXId(Query.FieldByName('sIdPlataforma').asstring)]
          else
            Grid.Cells[CsidPlataforma, i] := '';
          if trim(Query.FieldByName('sNumeroOrden').asstring) <> '@' then
            Grid.Cells[CsNumeroOrden, i] := Query.FieldByName('sNumeroOrden').asstring
          else
            Grid.Cells[CsNumeroOrden, i] := 'NA';
          Grid.Cell[CdCantidad, i].AsFloat := Query.FieldByName('dCantidad').AsFloat;
          Grid.Cells[CsAgrupaPersonal, i] := Query.FieldByName('sAgrupaPersonal').asstring;
          Grid.Cell[CdCantHH, i].AsFloat := Query.FieldByName('dCantHH').AsFloat;
          Grid.Cell[CdSolicitado, i].AsFloat := Query.FieldByName('dSolicitado').AsFloat;
          Grid.Cells[CsHoraInicio, i] := Query.FieldByName('sHoraInicio').asstring;
          Grid.Cells[CsHoraFinal, i] := Query.FieldByName('sHoraFinal').asstring;
          if LocalizaXId(Query.FieldByName('sTipoPernocta').asstring) <> -1 then
            Grid.Cells[CsTipoPernocta, i] := ClsTipoPErnocta.Items[LocalizaXId(Query.FieldByName('sTipoPernocta').asstring)]
          else
            Grid.Cells[CsTipoPernocta, i] := '';

          Grid.Cells[14, i] := Query.FieldByName('lImprimeResumen').asstring;

        end
        else
        begin
          Grid.Cells[CProcesar, i] := 'False';
          Grid.Cells[CsTipoObra, i] := '';
          Grid.Cells[CsIdPernocta, i] := ClsIdPernocta.items[0];
          Grid.Cells[CsIdPlataforma, i] := ClsIdPlataforma.items[0];
          Grid.Cells[CsNumeroOrden, i] := 'NA';
          Grid.Cell[CdCantidad, i].AsFloat := 0;
          Grid.Cells[CsAgrupaPersonal, i] := '';
          Grid.Cell[CdCantHH, i].AsFloat := 0;
          Grid.Cell[CdSolicitado, i].AsFloat := 0;
          Grid.Cells[CsHoraInicio, i] := '00:00';
          Grid.Cells[CsHoraFinal, i] := '00:00';
          Grid.Cells[CsTipoPernocta, i] := ClsTipoPernocta.items[0];
          Grid.Cells[14, i] := 'No';
        end;
      end;

      if (ComboBox.ItemIndex > -1) then
      begin
        EditAjustaPernocta.Enabled := True;
        Query.SQL.Text := 'SELECT * FROM bitacoradepernocta WHERE dIdFecha = :Fecha AND sContrato = :Contrato AND sNumeroOrden = :Folio ';
        Query.Params.ParamByName('Contrato').AsString := Global_Contrato;
        Query.Params.ParamByName('Fecha').AsString := FormatDateTime('yyyy-mm-dd', Fecha);
        Query.Params.ParamByName('Folio').AsString := Orden;
        Query.Open;
        EditAjustaPernocta.Value := Query.FieldByName('dCantidad').AsFloat;
      end
      else
      begin
        EditAjustaPernocta.Value := 0;
        EditAjustaPernocta.Enabled := False;
      end;
      
    finally
      if Orden = '@' then
        Orden := 'NA';
      lTITULOGRID.caption := 'RESUMEN DE PERSONAL DEL DIA: '+datetostr(fecha)+' PERTENECIENTE A LA CATEGORIA: '+Categoria+' EN FOLIO:  '+Orden;
      screen.Cursor := CurCantidades;
      Grid.CalculateFooter();
      Query.free;
    end;
  except
    on e:exception do
      showmessage('Ocurrió el siguiente error al cargar las cantidades de personal: '+#10+e.message);
  end;
end;

procedure TFrmResumenPersonal.GuardarCambios(Grid:TnextGrid;Categoria,Folio:String;Fecha:TdateTime);
Var
  Query: TZQuery;
  i,id:integer;
  CurSave:Tcursor;

  sOpcionLocal: string;
  sDiferencia : string;
  dFactor : Double;
  sMinutos, sHoras : string;
begin
  Try
    if Folio = 'NA' then
      Folio := '@';
    Query := TZQuery.Create(nil);
    CurSave := screen.Cursor;
    try
      screen.cursor := crappstart;
      Query.Connection := Connection.zConnection;
      pBarraAvance.Max := NextGrid.RowCount;
      query.Connection.StartTransaction;
          Query.Active := False;
          //Buscamos el registro q apunta el grid y lo eliminamos
          Query.SQL.Text := 'DELETE FROM bitacoradepersonal_cuadre WHERE sContrato = :sContrato AND (:sNumeroOrden = "<TODOS>" OR (:sNumeroOrden <> "<TODOS>" AND sNumeroOrden = :sNumeroOrden)) AND dIdFecha = :dIdFecha  AND sTipoObra = :sTipoObra ';
          Query.ParamByName('sContrato').AsString := Global_Contrato;
          Query.ParamByName('dIdFecha').AsString := FormatDateTime('yyyy-mm-dd', Fecha);
          //Query.ParamByName('sIdPersonal').AsString := NextGrid.Cells[CsidPersonal, i];
          Query.ParamByName('sTipoObra').AsString := Categoria;
          Query.ParamByName('sNumeroOrden').AsString := Folio;
          Query.ExecSQL;
      for i := 0 to Grid.RowCount - 1 do
      begin

        //query.Connection.StartTransaction;
        try
          pBarraAvance.Position := pBarraAvance.Position + 1;
         { Query.Active := False;
          //Buscamos el registro q apunta el grid y lo eliminamos
          Query.SQL.Text := 'DELETE FROM bitacoradepersonal_cuadre WHERE sContrato = :sContrato AND (:sNumeroOrden = "<TODOS>" OR (:sNumeroOrden <> "<TODOS>" AND sNumeroOrden = :sNumeroOrden)) AND dIdFecha = :dIdFecha AND sIdPersonal = :sIdPersonal AND sTipoObra = :sTipoObra LIMIT 1';
          Query.ParamByName('sContrato').AsString := Global_Contrato;
          Query.ParamByName('dIdFecha').AsString := FormatDateTime('yyyy-mm-dd', Fecha);
          Query.ParamByName('sIdPersonal').AsString := NextGrid.Cells[CsidPersonal, i];
          Query.ParamByName('sTipoObra').AsString := Categoria;
          Query.ParamByName('sNumeroOrden').AsString := Folio;
          Query.ExecSQL;  }

          //Se procede a guardar los registros en la tabla bitacoradepersonal_cuadre
          if (NextGrid.Cells[CProcesar, i] = 'True') {and (NextGrid.Cell[CdCantidad, i].asfloat > 0)} then
          begin
            //:='No';
            Query.Active := False;
            Query.SQL.Clear;
            Query.SQL.Add( '' +
                            'INSERT IGNORE INTO bitacoradepersonal_cuadre ' +
                            ' ( sContrato,  sNumeroOrden,  dIdFecha,  sIdPersonal,sTipoObra,  sDescripcion,  sIdPernocta,  sIdPlataforma,  sHoraInicio,  sHoraFinal,  dCantidad, iitemorden,  sagrupapersonal,  laplicapernocta, dCantHH, dSolicitado,sTipoPernocta,lImprimeResumen)' +
                            ' VALUES ' +
                            ' (:sContrato, :sNumeroOrden, :dIdFecha, :sIdPersonal,:sTipoObra, :sDescripcion, :sIdPernocta, :sIdPlataforma, :sHoraInicio, :sHoraFinal, :dCantidad,:iitemorden, :sagrupapersonal, :laplicapernocta, :dCantHH, :dSolicitado, :sTipoPernocta,:lImprimeResumen); ' +
                            '');
            Query.ParamByName('sContrato').AsString := Global_Contrato;
            if trim(NextGrid.Cells[CsNumeroOrden, i]) <> 'NA' then
              Query.ParamByName('sNumeroOrden').AsString :=  NextGrid.Cells[CsNumeroOrden, i]
            else
            begin
              if Folio <> '<NA>' then
                Query.ParamByName('sNumeroOrden').AsString :=  Folio
              else
                Query.ParamByName('sNumeroOrden').AsString := '@';
            end;
            Query.ParamByName('dIdFecha').AsString := FormatDateTime('yyyy-mm-dd', Fecha);
            Query.ParamByName('sIdPersonal').AsString := NextGrid.Cells[CsIdPersonal, i];
            Query.ParamByName('sTipoObra').AsString := Categoria;
            Query.ParamByName('sDescripcion').AsString := NextGrid.Cells[CsDescripcion, i];

            //se carga el sidpernocta y el sidplataroama de acuerdo al folio
            connection.QryBusca.Active := false;
            connection.qrybusca.SQL.Clear;
            connection.QryBusca.sql.Text := ''+
                       ' select * from bitacoradepersonal where dIdFecha = :dIdFecha and sContrato = :scontrato '+
                       ' and sNumeroOrden = :sNumeroOrden limit 1 ';
            connection.QryBusca.ParamByName('dIdFecha').AsString := FormatDateTime('yyyy-mm-dd', Fecha);
            connection.QryBusca.ParamByName('scontrato').AsString := Global_Contrato;
            connection.QryBusca.ParamByName('sNumeroOrden').AsString := Query.ParamByName('sNumeroOrden').AsString;
            connection.QryBusca.Open;
            {
            if trim(NextGrid.Cells[Csidpernocta, i]) <> '' then
              Query.ParamByName('sIdPernocta').AsString := NextGrid.Cells[Csidpernocta, i]
            else
              Query.ParamByName('sIdPernocta').AsString := '@';

            if trim(NextGrid.Cells[CsidPlataforma, i]) <> '' then
              Query.ParamByName('sIdPlataforma').AsString := NextGrid.Cells[CsidPlataforma, i]
            else
              Query.ParamByName('sIdPlataforma').AsString := '@';
                                          }
            if connection.QryBusca.recordcount = 1 then
            begin
              Query.ParamByName('sIdPernocta').AsString := connection.QryBusca.FieldByName('sIdPernocta').asstring;
              Query.ParamByName('sIdPlataforma').AsString := connection.QryBusca.FieldByName('sIdPlataforma').asstring;
            end
            else
            begin
              Query.ParamByName('sIdPernocta').AsString := '@';
              Query.ParamByName('sIdPlataforma').AsString := '@';
            end;
            
            Query.ParamByName('sHoraInicio').AsString :=  NextGrid.Cells[CsHoraInicio, i];
            Query.ParamByName('sHoraFinal').AsString :=  NextGrid.Cells[CsHoraFinal, i];
            Query.ParamByName('dCantidad').AsFloat := NextGrid.Cell[CdCantidad, i].AsFloat;
            Query.ParamByName('iitemorden').AsString := '0';
            if trim(NextGrid.Cells[CsAgrupaPersonal, i]) <> '' then
              Query.ParamByName('sagrupapersonal').AsString := NextGrid.Cells[CsAgrupaPersonal, i]
            else
              Query.ParamByName('sagrupapersonal').AsString := '@';
            Query.ParamByName('laplicapernocta').AsString := 'Si';

            //Estableciendo el sidtipopernocta
            id := ClsTipoPErnocta.Items.IndexOf(NextGrid.Cells[CsTipoPernocta, i]);
            if id = -1 then
              Query.ParamByName('sTipoPernocta').AsString := '@';

            if id = -1 then
              id := ClsTipoPErnocta.Items.IndexOf(' '+NextGrid.Cells[CsTipoPernocta, i]);

            if id > -1 then
              if (assigned(ClsTipoPErnocta.Items.Objects[id])) and (ClsTipoPErnocta.Items.Objects[id] <> nil) then
                Query.ParamByName('sTipoPernocta').AsString := TtipoPernocta(ClsTipoPErnocta.Items.Objects[id]).Identificador;

            //Estableciendo el sidpernocta
            id := ClsidPErnocta.Items.IndexOf(NextGrid.Cells[CsidPernocta, i]);
            if id = -1 then
              Query.ParamByName('sidPernocta').AsString := '@';

            if id = -1 then
              id := ClsidPErnocta.Items.IndexOf(' '+NextGrid.Cells[CsidPernocta, i]);

            if id > -1 then
              if (assigned(ClsidPErnocta.Items.Objects[id])) and (ClsidPErnocta.Items.Objects[id] <> nil) then
                Query.ParamByName('sidPernocta').AsString := TPernocta(ClsidPErnocta.Items.Objects[id]).Identificador;

            {//Estableciendo la plataforma
            id := ClsidPlataforma.Items.IndexOf(NextGrid.Cells[CsidPlataforma, i]);
            if id = -1 then
              Query.ParamByName('sidplataforma').AsString := '@';
                                                                   
            if id = -1 then
              id := ClsidPlataforma.Items.IndexOf(' '+NextGrid.Cells[CsidPlataforma, i]);

            if id > -1 then
              if (assigned(ClsidPlataforma.Items.Objects[id])) and (ClsidPlataforma.Items.Objects[id] <> nil) then
                Query.ParamByName('sidplataforma').AsString := TPlataforma(Clsidplataforma.Items.Objects[id]).Identificador;
                                         }

            sDiferencia := sfnRestaHoras(Query.ParamByName('sHoraFinal').asstring,Query.ParamByName('sHoraInicio').asstring);

            sHoras   := Copy(sDiferencia,1,2);
            sMinutos := Copy(sDiferencia,4,2);
            dFactor  := ((strTofloat(sHoras)*60) + strTofloat(sMinutos)) / 1440;
            Query.ParamByName('dCantHH').AsFloat := (dFactor * Query.ParamByName('dCantidad').asfloat) * 2;

            //Query.ParamByName('dCantHH').AsFloat := 0;
            Query.ParamByName('dSolicitado').AsFloat := NextGrid.Cell[CdSolicitado, i].AsFloat;
            Query.ParamByName('lImprimeResumen').AsString := NextGrid.Cell[14, i].AsString;
            Query.ExecSQL;
          end;

          //Eliminar de la bitacora de pernocta
          Query.SQL.Text := 'DELETE FROM bitacoradepernocta WHERE dIdFecha = :Fecha AND sContrato = :Contrato AND sNumeroOrden = :Folio '+
                            'and sIdCuenta=""';
          Query.Params.ParamByName('Contrato').AsString := Global_Contrato;
          Query.Params.ParamByName('Fecha').AsString := FormatDateTime('yyyy-mm-dd', fecha);
          Query.Params.ParamByName('Folio').AsString :=Folio;
          Query.ExecSQL;

          Query.SQL.Text := 'INSERT INTO bitacoradepernocta (sContrato, dIdFecha, sNumeroOrden, dCantidad) ' +
                            'VALUES (:Contrato, :Fecha, :Folio, :Cantidad) ';
          Query.Params.ParamByName('Contrato').AsString := Global_Contrato;
          Query.Params.ParamByName('Fecha').AsString := FormatDateTime('yyyy-mm-dd', fecha);
          Query.Params.ParamByName('Folio').AsString := Folio;
          Query.Params.ParamByName('Cantidad').AsFloat := EditAjustaPErnocta.Value;
          Query.ExecSQL;
          
         // Query.connection.Commit;
        except
          on e:exception do
          begin
            //Query.connection.Rollback;
            raise;
          end;
        end;

      end;
      Query.connection.Commit;
      CargarCantidades(Grid,Fecha,Categoria,folio);
    finally
      pbarraavance.Position := 0;
      screen.cursor := CurSave;
      query.Free;
    end;
  except
    on e:exception do
    begin
      Query.connection.Rollback;
      showmessage('Ocurrió el siguiente error al tratar de guardar los registros: '+#10+e.message);
    end
  End;
end;

Procedure TFrmResumenPersonal.ImportarPersonal(Desde:TTabla;Modo:TModoImportacion;Fecha:Tdatetime);
var QOrigen,QAux: tzReadonlyquery;
    QDestino,QElimina: TZQuery;
    CurImp:Tcursor;

  sDiferencia : string;
  dFactor : Double;
  sMinutos, sHoras : string;

  Continuar:Boolean;
  ListaFolios:string;
begin
  //localizamos folios atacados en la fecha
  connection.QryBusca.Active := False;
  connection.QryBusca.SQL.Text := 'select group_concat(distinct(snumeroorden)) as lista from bitacoradeactividades where didfecha = :fecha and scontrato = :contrato group by sContrato';
  connection.QryBusca.ParamByName('contrato').AsString := global_contrato;
  connection.QryBusca.ParamByName('fecha').asstring :=  FormatDateTime('yyyy-mm-dd', Fecha);;
  connection.QryBusca.Open;
  try
    if connection.QryBusca.RecordCount > 0 then
      ListaFolios := connection.QryBusca.FieldByName('lista').AsString
    else
      raise exception.Create('No hay folios reportados a la fecha, porfavor revisar el personal cargado en esta fecha en su cuadre porfavor');

    ListaFolios := AnsiReplaceStr(ListaFolios,',','","');
    ListaFolios := '"'+ListaFolios+'"';
    QOrigen := TzReadonlyQuery.create(nil);
    CurImp := screen.cursor;
    try
      screen.Cursor := crappstart;
      Qorigen.connection := Connection.zConnection;
      QDestino := TzQuery.Create(nil);
      try
        QDestino.Connection := Connection.zConnection;
        QAux := TzReadonlyQuery.create(nil);
        try
          QAux.Connection := Connection.zConnection;
          Qorigen.Active := False;
          QDestino.Active := False;
          QAux.Active := False;
          QOrigen.SQL.clear;
          QDestino.SQL.clear;
          QAux.SQL.clear;
          //De acuerdo a la opcion seleccionada se debe cargar los datos desde el origen de la
          //importación
          case Desde of
            BPersonal:
            begin
              QOrigen.SQL.text := 'Select * from bitacoradepersonal where scontrato = :scontrato and didfecha = :didfecha';
             // QOrigen.SQL.text := 'Select sContrato,dIdFecha,sIdPersonal,iItemOrden,sTipoObra,sDescripcion,sIdPernocta,sIdPlataforma,sNumeroOrden,'+
             // ' min(sHoraInicio) as sHoraInicio, max(sHoraFinal)as sHoraFinal,sum(dCantidad) as dCantidad,sAgrupaPersonal,lAplicaPernocta,sTipoPernocta,sum(dCantHH) as dCantHH,dSolicitado from bitacoradepersonal '+
              //' where scontrato = :scontrato and didfecha = :didfecha group by scontrato,snumeroorden,sidpersonal';
              QOrigen.ParamByName('didfecha').AsString := FormatDateTime('yyyy-mm-dd', Fecha);
            end;
            BPersonalCuadre:
            begin
              //localizamos de la bitacora de personal_cuadre todo personal con folios resultante de la consulta anterior
              QOrigen.SQL.text := 'Select * from bitacoradepersonal_cuadre where scontrato = :scontrato and didfecha = :didfecha and snumeroorden in ('+listaFolios+')';
              QOrigen.ParamByName('didfecha').AsString := FormatDateTime('yyyy-mm-dd', Fecha-1);
            end;
          end;
          QOrigen.ParamByName('scontrato').AsString := global_contrato;
          QOrigen.open;

          if QOrigen.recordcount = 0 then
            raise exception.create('No hay personal cargado en esa fecha.');

          if Modo = Sustituir then
          begin
            Qdestino.SQL.Text := 'DELETE FROM bitacoradepersonal_cuadre WHERE sContrato = :sContrato AND dIdFecha = :dIdFecha';
            Qdestino.ParamByName('sContrato').AsString := Global_Contrato;
            Qdestino.ParamByName('dIdFecha').AsString := FormatDateTime('yyyy-mm-dd', Fecha);
            Qdestino.ExecSQL;
            Qdestino.Active := False;
            Qdestino.SQL.clear;
          end;

          //La query q inserta los registros en la tabla bitacoradepersonal_cuadre
          Qdestino.SQL.Text :=  'INSERT IGNORE INTO bitacoradepersonal_cuadre ' +
                                ' ( sContrato,sNumeroOrden,dIdFecha,sIdPersonal,sTipoObra,sDescripcion,sIdPernocta,sIdPlataforma,sHoraInicio,sHoraFinal,dCantidad,iitemorden,sagrupapersonal,laplicapernocta,dCantHH,dSolicitado,sTipoPernocta)' +
                                ' VALUES ' +
                                ' (:sContrato,:sNumeroOrden,:dIdFecha,:sIdPersonal,:sTipoObra,:sDescripcion,:sIdPernocta,:sIdPlataforma,:sHoraInicio,:sHoraFinal,:dCantidad,:iitemorden,:sagrupapersonal,:laplicapernocta,:dCantHH,:dSolicitado,:sTipoPernocta);' ;

          //Buscamos el registro a ver si existe
          QAux.sql.clear;
          QAux.SQL.Text := 'SELECT * FROM bitacoradepersonal_cuadre WHERE sContrato = :sContrato AND '+
                           '(:sNumeroOrden = "<TODOS>" OR (:sNumeroOrden <> "<TODOS>" AND sNumeroOrden = :sNumeroOrden)) '+
                           'AND dIdFecha = :dIdFecha AND sIdPersonal = :sIdPersonal AND sTipoObra = :sTipoObra LIMIT 1';
          QAux.ParamByName('sContrato').AsString := Global_Contrato;
          QAux.ParamByName('dIdFecha').AsString := FormatDateTime('yyyy-mm-dd', Fecha);

          Qorigen.first;
          Pbarraavance.Max := Qorigen.RecordCount;
          PbarraAvance.Position := 0;
          while not Qorigen.Eof do
          begin
            PbarraAvance.Position := Qorigen.recno +1;
            Continuar := False;
            //Revisamos que el personal se encuentre en el moe
            if QOrigen.FieldByName('stipoobra').AsString = 'PU' then
            begin
              connection.QryBusca.Active := False;
              connection.QryBusca.SQL.Clear;
              connection.QryBusca.SQL.Text := ''+
                      'SELECT ' +
                      ' p.sDescripcion AS sCategoria, ' +
                      ' p.sIdPersonal, ' +
                      ' p.sDescripcion ' +
                      'FROM moe AS m ' +
                      '	INNER JOIN moerecursos AS mr ' +
                      '		ON (mr.iidMoe=m.iidMoe) ' +
                      '	INNER JOIN personal AS p ' +
                      '		ON (p.scontrato=:Contrato AND p.sidpersonal=mr.sidRecurso) ' +
                      'INNER JOIN gruposdepersonal AS gp '+
                      '	ON (gp.iId = p.iId_AgrupadorPersonal) '+
                      'WHERE ' +
                      '	m.didfecha = (SELECT max(didfecha) FROM moe WHERE didfecha <= :Fecha AND sContrato = :ContratoNormal) ' +
                      ' AND m.sContrato = :ContratoNormal ' +
                      '	AND mr.eTipoRecurso = "Personal" ' +
                      '	AND p.sidpersonal = :idPersonal ' +
                      'GROUP BY gp.sGrupo '+
                      ' ORDER BY p.iItemOrden ASC';
              connection.QryBusca.ParamByName('Contrato').AsString := Global_Contrato_Barco;
              connection.QryBusca.ParamByName('ContratoNormal').AsString := Global_Contrato;
              connection.QryBusca.ParamByName('Fecha').AsString := FormatDateTime('yyyy-mm-dd', Fecha);
              connection.QryBusca.ParamByName('idPersonal').AsString := Qorigen.Fieldbyname('sIdPersonal').AsString;
              Connection.QryBusca.Open;
              if Connection.QryBusca.RecordCount <> 0 then
                Continuar := True;
            end
            else
              Continuar := True;  // se inserta el personal simpre y cuando este en el moe
            if Continuar then
            begin
              //Simplemente se setean los parametros
              Qdestino.ParamByName('sContrato').AsString := Global_Contrato;
              Qdestino.ParamByName('sNumeroOrden').AsString := Qorigen.Fieldbyname('sNumeroOrden').AsString;
              if Desde = BPersonal then
                Qdestino.ParamByName('dIdFecha').asdatetime := Qorigen.Fieldbyname('dIdFecha').asdatetime
              else
                Qdestino.ParamByName('dIdFecha').asstring :=  FormatDateTime('yyyy-mm-dd', Fecha);
              Qdestino.ParamByName('sIdPersonal').AsString := Qorigen.Fieldbyname('sIdPersonal').AsString;
              Qdestino.ParamByName('sTipoObra').AsString := Qorigen.Fieldbyname('sTipoObra').AsString;
              Qdestino.ParamByName('sDescripcion').AsString := Qorigen.Fieldbyname('sDescripcion').AsString;
              Qdestino.ParamByName('sIdPlataforma').AsString := Qorigen.Fieldbyname('sIdPlataforma').AsString;
              Qdestino.ParamByName('sIdPernocta').AsString := Qorigen.Fieldbyname('sIdPernocta').AsString;
              Qdestino.ParamByName('sHoraInicio').AsString := Qorigen.Fieldbyname('sHoraInicio').AsString;
              Qdestino.ParamByName('sHoraFinal').AsString := Qorigen.Fieldbyname('sHoraFinal').AsString;
              Qdestino.ParamByName('dCantidad').AsFloat := Qorigen.Fieldbyname('dCantidad').AsFloat;
              Qdestino.ParamByName('iitemorden').AsString :=Qorigen.Fieldbyname('iitemorden').AsString;
              Qdestino.ParamByName('sagrupapersonal').AsString :=Qorigen.Fieldbyname('sagrupapersonal').AsString ;
              Qdestino.ParamByName('laplicapernocta').AsString := Qorigen.Fieldbyname('laplicapernocta').AsString;
              Qdestino.ParamByName('dCantHH').AsFloat := Qorigen.Fieldbyname('dCantHH').AsFloat;
              //Qdestino.ParamByName('dCantHH').AsFloat := 0;
              Qdestino.ParamByName('dSolicitado').AsFloat := Qorigen.Fieldbyname('dSolicitado').AsFloat;
              Qdestino.ParamByName('sTipoPernocta').AsString := Qorigen.Fieldbyname('sTipoPernocta').AsString;
              case Modo of
                Sustituir:
                begin
                  //si se debe sustituir entonces ejecutamos la query
                  Qdestino.ExecSQL;
                end;
                Sumar:
                begin
                  //si se deben sumar cantidades entonces buscar el registro
                  QAux.Active := False;
                  QAux.ParamByName('sIdPersonal').AsString := Qorigen.Fieldbyname('sIdPersonal').AsString;
                  QAux.ParamByName('sTipoObra').AsString := Qorigen.Fieldbyname('sTipoObra').AsString;
                  QAux.ParamByName('sNumeroOrden').AsString := Qorigen.Fieldbyname('sNumeroOrden').AsString;
                  QAux.Open;

                  if QAux.recordcount > 0 then
                  begin
                    //si se encuentra debemos sumar las cantidades al parametro ya seteado
                    Qdestino.ParamByName('dCantidad').AsFloat := Qdestino.ParamByName('dCantidad').AsFloat+QAux.FieldByName('Dcantidad').asfloat;

                    sDiferencia := sfnRestaHoras(Qdestino.ParamByName('sHoraFinal').asstring, Qdestino.ParamByName('sHoraInicio').asstring);
                    sHoras   := Copy(sDiferencia,1,2);
                    sMinutos := Copy(sDiferencia,4,2);
                    dFactor  := ((strTofloat(sHoras)*60) + strTofloat(sMinutos)) / 1440;
                    Qdestino.ParamByName('dCantHH').AsFloat := (dFactor * Qdestino.ParamByName('dCantidad').AsFloat) * 2;


                    //antes de enviar el registro lo eliminamos
                    QElimina := tzquery.create(nil);
                    try
                      Qelimina.Connection := Connection.zConnection;
                      Qelimina.active := False;
                      Qelimina.sql.clear;
                      Qelimina.SQL.Text := 'Delete FROM bitacoradepersonal_cuadre WHERE sContrato = :sContrato AND '+
                                       '(:sNumeroOrden = "<TODOS>" OR (:sNumeroOrden <> "<TODOS>" AND sNumeroOrden = :sNumeroOrden)) '+
                                       'AND dIdFecha = :dIdFecha AND sIdPersonal = :sIdPersonal AND sTipoObra = :sTipoObra LIMIT 1';
                      Qelimina.ParamByName('sContrato').AsString := Global_Contrato;
                      Qelimina.ParamByName('dIdFecha').AsString := FormatDateTime('yyyy-mm-dd', Fecha);
                      Qelimina.ParamByName('sIdPersonal').AsString := Qorigen.Fieldbyname('sIdPersonal').AsString;
                      Qelimina.ParamByName('sTipoObra').AsString := Qorigen.Fieldbyname('sTipoObra').AsString;
                      Qelimina.ParamByName('sNumeroOrden').AsString := Qorigen.Fieldbyname('sNumeroOrden').AsString;
                      Qelimina.ExecSQL;
                    finally
                      Qelimina.Free;
                    end;
                    Qdestino.ExecSQL;
                  end
                  else
                  begin
                    //En caso de no encontrarse entonces se ejecuta la query
                    Qdestino.ExecSQL;
                  end;

                end;
                Agregar:
                begin
                  //buscar el registro
                  QAux.Active := False;
                  QAux.ParamByName('sIdPersonal').AsString := Qorigen.Fieldbyname('sIdPersonal').AsString;
                  QAux.ParamByName('sTipoObra').AsString := Qorigen.Fieldbyname('sTipoObra').AsString;
                  QAux.ParamByName('sNumeroOrden').AsString := Qorigen.Fieldbyname('sNumeroOrden').AsString;
                  QAux.Open;

                  if QAux.recordcount = 0 then
                  begin
                    //si no se encuentra entonces ejecutar la query
                    Qdestino.ExecSQL;
                  end;
                end;
              end;
            end;

            QOrigen.Next;
          end;
        finally
          Qaux.free;
        end;
      finally
        QDestino.Free;
      end;
    finally
      PbarraAvance.Position := 0;
      screen.Cursor := CurImp;
      QOrigen.free;
    end;

  except
    on e:Exception do
      showmessage('Ocurrió el siguiente error al importar personal: '+#10+e.message)
  end;

end;

procedure TFrmResumenPersonal.Imppersonaldediaanteriorcondiferentefolio1Click(
  Sender: TObject);
var
  TzFolioA:TZreadonlyquery;
  rPAdre :Integer;
  ufolio:string;
  fol:TFolio;
begin
  TzFolioA := TZreadonlyquery.Create(nil);
  try
    TzFolioA.Connection := connection.zConnection;
    TzFolioA.Active := False;
    TzFolioA.SQL.Clear;                                                                                             {and sNumeroOrden <> "@"}
    TzFolioA.SQL.Text := 'select * from bitacoradepersonal_cuadre where scontrato = :contrato and dIdFecha = :fecha   order by sNumeroOrden';
    TzFolioA.ParamByName('contrato').AsString := global_contrato;
    TzFolioA.ParamByName('fecha').AsDate := dIdFecha.Date-1;
    TzFolioA.Open;

    ufolio := '';
    TzFolioA.First;
    GrdPersonal.ClearRows;
    while not TzFolioA.Eof do
    begin
      if (ufolio <> TzFolioA.FieldByName('snumeroorden').AsString) then
      begin
        GrdPersonal.AddRow(1);
        GrdPersonal.Cells[1,GrdPersonal.RowCount-1] := TzFolioA.FieldByName('snumeroorden').AsString;
        rPAdre := GrdPersonal.RowCount-1;
        ufolio := TzFolioA.FieldByName('snumeroorden').AsString;
      end
      else
      begin
        GrdPersonal.AddChildRow(rpadre);
        GrdPersonal.Cells[3,GrdPersonal.RowCount-1] := TzFolioA.FieldByName('dcantidad').AsString;
        GrdPersonal.Cells[4,GrdPersonal.RowCount-1] := TzFolioA.FieldByName('sidpersonal').AsString;
        GrdPersonal.Cells[5,GrdPersonal.RowCount-1] := TzFolioA.FieldByName('stipoobra').AsString;
        GrdPersonal.Cells[6,GrdPersonal.RowCount-1] := TzFolioA.FieldByName('sdescripcion').AsString;
        GrdPersonal.Cells[7,GrdPersonal.RowCount-1] := TzFolioA.FieldByName('stipopernocta').AsString;

        TzFolioA.Next;
      end;

    end;
  finally
    TzFolioA.Free;
  end;

  TzFolioA := TZreadonlyquery.Create(nil);
  try
    TzFolioA.Connection := connection.zConnection;
    TzFolioA.Active := False;
    TzFolioA.SQL.Clear;
    TzFolioA.SQL.Text :=  'select ot.sNumeroOrden,ot.sIdPlataforma,ot.sIdPernocta from bitacoradeactividades ba '+
                          'inner join ordenesdetrabajo ot '+
                          'on (ot.scontrato = ba.sContrato and ot.sNumeroOrden = ba.sNumeroOrden) '+
                          'where ba.sContrato = :Contrato and ba.dIdFecha = :fecha group by ba.sNumeroOrden ';
    TzFolioA.ParamByName('contrato').AsString := global_contrato;
    TzFolioA.ParamByName('fecha').AsDate := dIdFecha.Date;
    TzFolioA.Open;
    TzFolioA.First;
    cmbChanFolio.Items.Clear;
    while not TzFolioA.Eof do
    begin
      fol := TFolio.Create;
      fol.Identificador := TzFolioA.FieldByName('snumeroorden').AsString;
      fol.Plataforma := TzFolioA.FieldByName('sidplataforma').AsString;
      fol.Pernoctan := TzFolioA.FieldByName('sidpernocta').AsString;
      cmbChanFolio.AddItem(TzFolioA.FieldByName('snumeroorden').AsString,fol);
      TzFolioA.Next;
    end;
  finally
    TzFolioA.Free;
  end;
  if cmbChanFolio.Items.Count > 0 then
  begin
    cmbChanFolio.ItemIndex := 0;
    cmbChanFolioChange(cmbChanFolio);
  end;
  PnlDiaAnterior.Visible := True;
end;

Procedure TFrmResumenPersonal.DesSeleccionarTodo(Grid:TnextGrid;TP:Boolean);
var i:integer;
begin
  for i := 0 to Grid.rowCount - 1 do
  begin
    if tp then
      Grid.Cells[CProcesar,i] := 'true'
    else
    begin
      Grid.Cells[CProcesar,i] := 'false';
      Grid.Cells[7,i] := '0';
    end;
  end;
end;
{$ENDREGION}
end.
