unit frm_CuadreXCategoria;

interface

uses

  {$region 'Uses'}

  UnitCuadre, global, frm_connection, Math, Types, Messages, SysUtils,
  Variants, Classes, Graphics, Controls, Forms, ComObj,
  Dialogs, dxSkinsCore, dxSkinDevExpressStyle, cxSSTypes,
  dxSkinsdxDockControlPainter, cxClasses, dxDockControl, cxPC, dxDockPanel,
  dxSkinFoggy, cxGraphics, cxControls, cxLookAndFeels, cxLookAndFeelPainters,
  cxContainer, cxEdit, cxGroupBox, ComCtrls, cxListView, ImgList, cxTreeView,
  DB, ZAbstractRODataset, ZAbstractDataset, ZDataset, Menus,
  StdCtrls, cxButtons, dxtree, DBClient, cxTextEdit, cxMaskEdit, cxDropDownEdit,
  cxImageComboBox, cxSSheet, cxLabel, cxLookupEdit,
  cxDBLookupEdit, cxDBLookupComboBox, cxDBLabel, cxMemo, cxDBEdit, cxSpinEdit,
  cxTimeEdit, Mask, JvExMask, JvSpin, cxCalc, dxSkinOffice2010Blue, cxListBox ;

  {$endregion}

type
  TfrmCuadreCategoria = class(TForm)

    {$region 'Componentes y Procedimientos'}

    dxDockManager: TdxDockingManager;
    dxDock: TdxDockSite;
    dxDockPanel: TdxDockPanel;
    dxLayoutDockSite1: TdxLayoutDockSite;
    dxpnlCategorias: TdxDockPanel;
    dxPanelCuadre: TdxLayoutDockSite;
    cxGroupBox1: TcxGroupBox;
    treeCategorias: TcxTreeView;
    tsIconos: TcxImageList;
    btnGuardar: TcxButton;
    cxIconosBotones32: TcxImageList;
    cxImageList1: TcxImageList;
    grpDias: TcxGroupBox;
    treeCuadres: TcxTreeView;
    btnSeleccionarFecha: TcxButton;
    cxIconosBotones16: TcxImageList;
    cbbCategorias: TcxImageComboBox;
    cximgCombo: TcxImageList;
    cxButton2: TcxButton;
    cxCuadre: TcxSpreadSheet;
    grpCategoria: TcxGroupBox;
    prgFolios: TProgressBar;
    cxLabel1: TcxLabel;
    cxLabel2: TcxLabel;
    cxLabel3: TcxLabel;
    prgActividades: TProgressBar;
    prgHorarios: TProgressBar;
    grpDatosCategoria: TcxGroupBox;
    dbPlataforma: TcxDBLookupComboBox;
    dsCategorias: TDataSource;
    cxLabel4: TcxLabel;
    cxLabel5: TcxLabel;
    dbPernocta: TcxDBLookupComboBox;
    zPernoctas: TZQuery;
    zPlataformas: TZQuery;
    dsPernoctas: TDataSource;
    dsPlataformas: TDataSource;
    popPrincipal: TPopupMenu;
    btnpopCrearHorario: TMenuItem;
    cxGroupBox2: TcxGroupBox;
    cxLabel6: TcxLabel;
    dsActividades: TDataSource;
    dbActividadDescripcion: TcxDBMemo;
    dsFolios: TDataSource;
    grpCrearHorario: TcxGroupBox;
    lblCaption: TcxLabel;
    dsHorarios: TDataSource;
    btnAceptarHorario: TcxButton;
    btnCancelarHorario: TcxButton;
    cxHoraInicio: TcxTimeEdit;
    grpProgreso: TcxGroupBox;
    dbLblbCategoria: TcxDBLabel;
    cxLabel8: TcxLabel;
    cxDBLabel1: TcxDBLabel;
    dbSolicitado: TcxDBLabel;
    cxLabel9: TcxLabel;
    dbA_Bordo: TcxDBLabel;
    cxLabel10: TcxLabel;
    cxLabel11: TcxLabel;
    dbCantidad: TcxDBLabel;
    btnpopEditarHorario: TMenuItem;
    btnpopEliminarHorario: TMenuItem;
    ImprimirEstructura1: TMenuItem;
    cxHoraFinal: TcxMaskEdit;
    cdFolios: TClientDataSet;
    cdActividades: TClientDataSet;
    cdCategorias: TClientDataSet;
    cdHorarios: TClientDataSet;
    grpAjuste: TcxGroupBox;
    cxLabel12: TcxLabel;
    cxLabel13: TcxLabel;
    btnAplicaAjuste: TcxButton;
    Ajustes1: TMenuItem;
    lstCuentas: TcxListBox;
    clcAjuste: TcxCalcEdit;
    lblSuma: TcxLabel;
    procedure FormCreate(Sender: TObject);
    procedure CambioDeNodo(Sender: TObject; Node: TTreeNode);
    procedure cbbCategoriasPropertiesChange(Sender: TObject);
    procedure cxButton2Click(Sender: TObject);
    procedure CambioDeFecha(Sender: TObject; Node: TTreeNode);
    procedure cxCuadreSetSelection(Sender: TObject; ASheet: TcxSSBookSheet);
    procedure PrepararNuevoHorario(Sender: TObject);
    procedure btnGuardarClick(Sender: TObject);
    procedure cxCuadreClearCells(Sender: TcxSSBookSheet; const ACellRect: TRect;
      var UseDefaultStyle, CanClear: Boolean);
    procedure cxCuadreEndEdit(Sender: TObject);
    procedure cxCuadreKeyPress(Sender: TObject; var Key: Char);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure EditarHorario(Sender: TObject);
    procedure EliminarHorario(Sender: TObject);
    procedure ImprimirEstructura1Click(Sender: TObject);
    procedure Ajustes1Click(Sender: TObject);
    procedure btnAplicaAjusteClick(Sender: TObject);
    procedure lstCuentasClick(Sender: TObject);
    procedure clcAjustePropertiesChange(Sender: TObject);

    {$endregion}

  private
    { Private declarations }
    Formula : TftFormula;

    procedure ComprobarSuma( SumarValor : Double );
    function CategoriaBloqueada( Categoria : string ):Boolean;
    function ObtenerFilaNuevoHorario():TExcelRow;
  public
    { Public declarations }
  end;  

var
  frmCuadreCategoria: TfrmCuadreCategoria;

implementation

uses
  Windows;

{$R *.dfm}

procedure TfrmCuadreCategoria.Ajustes1Click(Sender: TObject);
var
  ZCuentas : TZReadOnlyQuery;

  Form : TForm;
begin
  try

    InicializarZQuery( ZCuentas );
    ZCuentas.SQL.Text := 'select sIdPernocta from cuentas';
    ZCuentas.Open;
    ZCuentas.First;

    lstCuentas.Items.Clear;
    while not ZCuentas.Eof do
    begin
      lstCuentas.Items.Add( ZCuentas.FieldByName( 'sIdPernocta' ).AsString );
      ZCuentas.Next;
    end;

    InicializarForm( Form, grpAjuste, 300, 200 );
    btnAplicaAjuste.Enabled := False;
    Form.ShowModal;

    with grpAjuste do
    begin
      Parent := Self;
      Align := alNone;
      Left := 0;
      Top := 0;
      Width := 0;
      Height := 0;
      Visible := False;
    end;

  finally
    Form.Free;
    ZCuentas.Free;
  end;
end;

procedure TfrmCuadreCategoria.btnGuardarClick(Sender: TObject);
var
  Ruta : TFileName;
begin

  Ruta := Cuadre.Path + Cuadre.GenerarNombreAleatorio() + '.xls';
  cxCuadre.SaveToFile( Ruta );

  Cuadre.GuardarCuadre( cdFolios, cdActividades, cdCategorias, cdHorarios, Ruta );
  cxCuadre.LoadFromFile( Ruta );
  
end;

procedure TfrmCuadreCategoria.CambioDeFecha(Sender: TObject; Node: TTreeNode);
begin
  Node.SelectedIndex := Node.ImageIndex;
  btnSeleccionarFecha.Enabled := Node.ImageIndex = 4;
end;

procedure TfrmCuadreCategoria.CambioDeNodo(Sender: TObject; Node: TTreeNode);
var
  Ruta : TFileName;
  Categoria : string;

  function ObtenerCategoria( Source : string ):string;
  var
    C : integer;
  begin
    for C := 1 to Length( Source ) - 1 do
    begin
      if Source[ C ] = '-' then
        Break;

      Result := Result + Trim( Source[ C ] );
    end;
  end;

begin

  try
    if Node.ImageIndex > 1 then
      Node.SelectedIndex := 5
    else
      Node.SelectedIndex := Node.ImageIndex;

    if Node.ImageIndex = 2 then
    begin

      if Cuadre.Cambios then
      begin
        if MessageDlg( 'Se han detectado cambios en el cuadre actual, si se cambia de categoria, se perderán estos cambios.'+#13+#10+' ¿Desea continuar?', mtInformation, [ mbYes, mbNo ], 0 ) = mrNo then
          raise Exception.Create( 'Proceso cancelado.' );
      end;

      //En caso de tener ya una categoria ocupada, el simple hecho de cambiar a otra
      //Hay que desbloquear la que ya tengamos sea cual sea.

      Categoria := ObtenerCategoria( treeCategorias.Selected.Text );

      if CategoriaBloqueada( Categoria ) then
        raise Exception.Create( 'Alguien ya esta ocupando esta categoria' );

      Cuadre.DesbloquearCategoria;
      Sleep( 150 );

      Cuadre.Categoria := Categoria;
      with connection.QryBusca do
      begin

        Active := False;
        if cbbCategorias.ItemIndex = 0 then
        begin

          
          SQL.Text := 'select '+
                        'replace( upper(p.sDescripcion), " ", "" ) as Descripcion, '+
                        'upper( a.sTierra ) as Tierra '+
                      'from personal as p '+
                      'inner join anexos as a '+
                        'on ( a.sAnexo = p.sAnexo and a.sTipo = "PERSONAL" ) '+

                      'where '+
                        'p.sContrato = :contrato and '+
                        'p.lCobro = "Si" and '+
                        'p.sIdPersonal = :idrecurso; ';

        end
        else
          SQL.Text := 'select '+
                        'replace( upper(e.sDescripcion), " ", "" ) as Descripcion '+
                      'from equipos as e '+
                      'where '+
                        'e.sContrato = :contrato and '+
                        'e.sIdEquipo = :idrecurso and '+
                        'e.lCobro = "Si"; ';
        begin
        end;

        ParamByName( 'contrato' ).AsString := Cuadre.Contrato;
        ParamByName( 'idrecurso' ).AsString := Cuadre.Categoria;
        Open;

        if FieldByName( 'Descripcion' ).AsString = 'TIEMPOEXTRA' then
          Formula := ftTiempoExtra
        else
        begin
          if cbbCategorias.Text = 'Personal' then
          begin
            if FieldByName( 'Tierra' ).AsString = 'SI' then
              Formula := ftTierra
            else
              Formula := ftA_Bordo;
          end
          else
            Formula := ftEq;

        end;

      end;
      
      
      Application.ProcessMessages;
      //Bloquear la categoria
      Cuadre.BloquearCategoria;
      Sleep( 350 );
      cxCuadre.Visible := False;
      dxpnlCategorias.Enabled := False;
      grpProgreso.Visible := True;
      dbActividadDescripcion.DataBinding.DataSource := nil;
      Application.ProcessMessages;

      if cdHorarios.RecordCount > 0 then
        cdHorarios.EmptyDataSet;

      Cuadre.CargarCategorias( cdCategorias, TftCategoria( cbbCategorias.ItemIndex ) );
      Cuadre.CargarFolios( cdFolios );

      if cdFolios.RecordCount = 0 then
        raise Exception.Create( 'No se encontrarón actividades reportadas.' );

      Cuadre.CargarActividades( cdActividades, cdFolios, Formula );
      Cuadre.CargarHorarios( cdFolios, cdActividades, cdHorarios, cdCategorias );
      Ruta := Cuadre.GenerarCuadre( cdFolios, cdCategorias, cdActividades, cdHorarios, prgFolios, prgActividades, prgHorarios, Formula );

      cxCuadre.LoadFromFile( Ruta );
      cxCuadre.Visible := True;
      grpProgreso.Visible := False;
      dxpnlCategorias.Enabled := True;
      Cuadre.Cambios := False;
      dbActividadDescripcion.DataBinding.DataSource := dsActividades;
      Application.ProcessMessages;

      if FileExists( Ruta ) then
        DeleteFile( PChar( Ruta ) );

    end;

  except
    on e:Exception do
      MessageDlg( e.Message, mtInformation, [ mbOK ], 0 );

  end;

end;

procedure TfrmCuadreCategoria.cbbCategoriasPropertiesChange(Sender: TObject);
var
  Tipo : TftCategoria;
begin

  Cuadre.DesbloquearCategoria;
  Tipo := cbbCategorias.Properties.Items[ cbbCategorias.SelectedItem ].Value;

  Cuadre.Tipo := 'Equipo';
  Cuadre.TipoCategoria := ftEquipo;
  if Tipo = ftPersonal then
  begin
    Cuadre.Tipo := 'Personal';
    Cuadre.TipoCategoria := ftPersonal;
  end;

  MostrarCategorias( treeCategorias, Tipo, False );

end;

procedure TfrmCuadreCategoria.clcAjustePropertiesChange(Sender: TObject);
begin
  btnAplicaAjuste.Enabled := True;
end;

procedure TfrmCuadreCategoria.PrepararNuevoHorario(Sender: TObject);
var
  IdHorario : TIdRegistroHorario;
  IdActividad : TIdRegistroFolio;

  Fila : TExcelRow;
  Form : TForm;

  Archivo : TFileName;

begin

  try

    Fila := cxCuadre.ActiveCell.Y + 1;
    if ( Cuadre.EncuentraFolio( cdFolios, Fila ) ) and ( Cuadre.EncuentraActividad( cdActividades, Fila ) ) or ( Cuadre.EncuentraHorario( cdHorarios, Fila ) ) then
    begin
      IdActividad := cdActividades.FieldByName( 'IdRegistro' ).AsInteger;
      InicializarForm( Form, grpCrearHorario, 240, 170 );
      cxHoraInicio.Text := cdActividades.FieldByName( 'HInicio' ).AsString;
      cxHoraFinal.Text := cdActividades.FieldByName( 'HFin' ).AsString;
      lblCaption.Caption := 'NUEVO HORARIO';
      if Form.ShowModal = mrOk then
      begin

        if cxHoraFinal.Text = '24:00' then
        begin
          if not ( ( StrToTime( '23:59' ) + StrToTime('00:01') ) > cxHoraInicio.Time ) then
            raise Exception.Create( 'La hora de inicio no puede ser mayor a la de termino' );
        end
        else
        begin
          if not  ( StrToTime( cxHoraFinal.Text ) > cxHoraInicio.Time ) then
            raise Exception.Create( 'La hora de inicio no puede ser mayor a la de termino' );
        end;

        if cxHoraInicio.Time < StrToTime( cdActividades.FieldByName( 'HInicio' ).AsString ) then
          raise Exception.Create( 'La hora de inicio del corte no puede ser menor a la hora inicial de la actividad' );

        if cdActividades.FieldByName( 'HFin' ).AsString = '24:00' then
        begin
          if cxHoraFinal.Text = '24:00' then
          begin
            if (( StrToTime( '23:59' ) + StrToTime('00:01') ) > ( StrToTime( '23:59' ) + StrToTime('00:01') ) ) then
              raise Exception.Create( 'La hora de termino del horario no puede ser mayor que la hora final de la actividad' );
          end
          else
          begin
            if ( StrToTime( cxHoraFinal.Text ) > ( StrToTime( '23:59' ) + StrToTime('00:01') ) ) then
              raise Exception.Create( 'La hora de termino del horario no puede ser mayor que la hora final de la actividad' );
          end;
        end
        else
        begin
          if ( StrToTime( cxHoraFinal.Text ) > StrToTime( cdActividades.FieldByName( 'HFin' ).AsString ) ) then
            raise Exception.Create( 'La hora de termino del horario no puede ser mayor que la hora final de la actividad' );
        end;

        ReestablecerBox( grpCrearHorario, Self );

        if Cuadre.EncuentraHorario( cdHorarios, Fila ) then
          Inc( Fila );

        Fila := ObtenerFilaNuevoHorario;
        IdHorario := Cuadre.CrearHorario( IdActividad, cdhorarios, cxHoraInicio.Text, cxHoraFinal.Text, Fila, Formula );
        Cuadre.ActualizarFilasConjuntoActual( cdFolios, cdActividades );
        Cuadre.ActualizarFilasFolios( cdFolios );
        Cuadre.ActualizarFilasActividades( cdActividades );
        Cuadre.ActualizarFilasHorarios( cdHorarios, 1, ftInsert, Fila );
        Archivo := Cuadre.Path + Cuadre.GenerarNombreAleatorio( ) + '.xls';
        cxCuadre.SaveToFile( Archivo );
        cdHorarios.Locate( 'IdRegistro', IdHorario, [] );
        cxCuadre.LoadFromFile( Cuadre.RegenerarCuadre( cdFolios, cdActividades, cdHorarios, Archivo, Fila, cxHoraInicio.Text, cxHoraFinal.Text, cdHorarios.FieldByName( 'Duracion' ).AsString, Formula ) );
        cdHorarios.Locate( 'IdRegistro', IdHorario, [] );

        Application.ProcessMessages;
        cxCuadre.Sheet.SelectCell( 5, cdHorarios.FieldByName( 'Fila' ).AsInteger - 1, False );

        if FileExists( Archivo ) then
          DeleteFile( PCHar( Archivo ) );

      end;

    end;

  except
    on e:Exception do
    begin
      MessageDlg( e.Message, mtInformation, [ mbOK ], 0 )
    end;

  end;

end;

procedure TfrmCuadreCategoria.btnAplicaAjusteClick(Sender: TObject);
var
  ZAjuste : TZQuery;
  IdCuenta : string;
begin

  try
    InicializarZQuery( ZAjuste );
    with ZAjuste do
    begin
      Active := False;
      SQL.Text := 'select '+
                    '(select count( sIdCuenta ) from bitacoradepernocta '+
                        'where '+
                          'sContrato = :orden and '+
                          'dIdFecha = :fecha and '+
                          'sNumeroOrden = :folio and '+
                          'sIdCuenta = (select sIdCuenta '+
                                       'from cuentas '+
                                       'where sIdPernocta = :pernocta ) limit 1 '+
                      ') as ExisteAjuste, '+
                    '(select sIdCuenta from cuentas where sIdPernocta = :pernocta ) as sIdCuenta;';
      ParamByName( 'orden' ).AsString := Cuadre.Orden;
      ParamByName( 'fecha' ).AsString := Cuadre.Fecha;
      ParamByName( 'folio' ).AsString := cdFolios.FieldByName( 'Folio' ).AsString;
      ParamByName( 'pernocta' ).AsString := lstCuentas.Items[ lstCuentas.ItemIndex ];
      Open;

      IdCuenta := FieldByName( 'sIdCuenta' ).AsString;
    end;

    if ZAjuste.FieldByName( 'ExisteAjuste' ).AsInteger > 0 then
    begin
      with ZAjuste do
      begin
        Active := False;
        SQL.Text := 'update bitacoradepernocta set dCantidad = :ajuste '+
                    'where '+
                      'sContrato = :orden and '+
                      'dIdFecha = :fecha and '+
                      'sNumeroOrden = :folio and '+
                      'sIdCuenta = :cuenta; ';

      end;
    end
    else
    begin
      with ZAjuste do
      begin
        Active := False;
        SQL.Text := 'insert into bitacoradepernocta '+
                    'values( :orden, :fecha, :folio, :cuenta, 0, :ajuste ); ';

      end;  
    end;

    ZAjuste.ParamByName( 'ajuste' ) .AsFloat := clcAjuste.Value;
    ZAjuste.ParamByName( 'orden' ).AsString := Cuadre.Orden;
    ZAjuste.ParamByName( 'fecha' ).AsString := Cuadre.Fecha;
    ZAjuste.ParamByName( 'folio' ).AsString := cdFolios.FieldByName( 'Folio' ).AsString;
    ZAjuste.ParamByName( 'cuenta' ).AsString := IdCuenta;
    ZAjuste.ExecSQL;

    btnAplicaAjuste.Enabled := False;
    
  finally
    ZAjuste.Free;
  end;

end;

procedure TfrmCuadreCategoria.cxButton2Click(Sender: TObject);
var
  zDias : TZReadOnlyQuery;
begin
  InicializarZQuery( zDias );
  zDias.SQL.Text := 'select dIdFecha from reportediario where sOrden = :orden';
  zDias.ParamByName( 'orden' ).AsString := Cuadre.Orden;
  zDias.Open;
  zDias.First;

  treeCuadres.Items.Clear;
  while not zDias.Eof do
  begin
    with treeCuadres.Items.AddChild( nil, zDias.FieldByName( 'dIdFecha' ).AsString ) do
      ImageIndex := 4;
    zDias.Next;
  end;

end;

procedure TfrmCuadreCategoria.cxCuadreClearCells(Sender: TcxSSBookSheet;
  const ACellRect: TRect; var UseDefaultStyle, CanClear: Boolean);
var
  Rango : TcxSSCellObject;
begin
  Rango := cxCuadre.Sheet.GetCellObject( 6, cxCuadre.ActiveCell.Y );

  try
    ComprobarSuma( -StrToFloat( VarToStr( Rango.CellValue ) ) );
  except
    ComprobarSuma( 0.0 );
  end;
end;

procedure TfrmCuadreCategoria.cxCuadreEndEdit(Sender: TObject);
begin
  ComprobarSuma( 0.0 );
end;

procedure TfrmCuadreCategoria.cxCuadreKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #8 then
    cxCuadre.Sheet.GetCellObject( cxCuadre.ActiveCell.X, cxCuadre.ActiveCell.Y ).Text := EMPTY_STRING;
end;

procedure TfrmCuadreCategoria.cxCuadreSetSelection(Sender: TObject;
  ASheet: TcxSSBookSheet);
begin
  Application.ProcessMessages;
  Cuadre.EncuentraFolio( cdFolios, ASheet.ActiveCell.Y + 1 );
  Application.ProcessMessages;
  if Cuadre.EncuentraActividad( cdActividades, ASheet.ActiveCell.Y + 1 ) then
    dxDockPanel.PopupMenu := popPrincipal
  else
    dxDockPanel.PopupMenu := nil;
end;

procedure TfrmCuadreCategoria.EditarHorario(Sender: TObject);
var
  Form : TForm;
  Fila : TExcelRow;
  IdHorario : TIdRegistroHorario;
  IdActividad : TIdRegistroFolio;

  Excel,
  Libro,
  Hoja,
  Rango : TExcelInstance;

  Ruta : string;

begin
  try
    Fila := cxCuadre.ActiveCell.Y + 1;
    if ( Cuadre.EncuentraFolio( cdFolios, Fila ) ) and ( Cuadre.EncuentraActividad( cdActividades, Fila ) ) or ( Cuadre.EncuentraHorario( cdHorarios, Fila ) ) then
    begin
      IdActividad := cdActividades.FieldByName( 'IdRegistro' ).AsInteger;
      InicializarForm( Form, grpCrearHorario, 240, 170 );
      cxHoraInicio.Text := cdHorarios.FieldByName( 'Inicio' ).AsString;
      cxHoraFinal.Text := cdHorarios.FieldByName( 'Fin' ).AsString;
      lblCaption.Caption := 'EDITAR HORARIO';

      if Form.ShowModal = mrOk then
      begin
        if cxHoraFinal.Text = '24:00' then
        begin
          if not ( ( StrToTime( '23:59' ) + StrToTime('00:01') ) > cxHoraInicio.Time ) then
            raise Exception.Create( 'La hora de inicio no puede ser mayor a la de termino' );
        end
        else
        begin
          if not  ( StrToTime( cxHoraFinal.Text ) > cxHoraInicio.Time ) then
            raise Exception.Create( 'La hora de inicio no puede ser mayor a la de termino' );
        end;

        if cxHoraInicio.Time < StrToTime( cdActividades.FieldByName( 'HInicio' ).AsString ) then
          raise Exception.Create( 'La hora de inicio del corte no puede ser menor a la hora inicial de la actividad' );

        if cdActividades.FieldByName( 'HFin' ).AsString = '24:00' then
        begin
          if cxHoraFinal.Text = '24:00' then
          begin
            if (( StrToTime( '23:59' ) + StrToTime('00:01') ) > ( StrToTime( '23:59' ) + StrToTime('00:01') ) ) then
              raise Exception.Create( 'La hora de termino del horario no puede ser mayor que la hora final de la actividad' );
          end
          else
          begin
            if ( StrToTime( cxHoraFinal.Text ) > ( StrToTime( '23:59' ) + StrToTime('00:01') ) ) then
              raise Exception.Create( 'La hora de termino del horario no puede ser mayor que la hora final de la actividad' );
          end;
        end
        else
        begin
          if ( StrToTime( cxHoraFinal.Text ) > StrToTime( cdActividades.FieldByName( 'HFin' ).AsString ) ) then
            raise Exception.Create( 'La hora de termino del horario no puede ser mayor que la hora final de la actividad' );
        end;

        ReestablecerBox( grpCrearHorario, Self );
        IdHorario := cdHorarios.FieldByName( 'IdRegistro' ).AsInteger;
        cdHorarios.Edit;
        cdHorarios.FieldByName( 'Inicio' ).AsString := cxHoraInicio.Text;
        cdHorarios.FieldByName( 'Fin' ).AsString := cxHoraFinal.Text;
        cdHorarios.FieldByName( 'Duracion' ).AsString := Cuadre.ObtenerDuracion( cxHoraFinal.Text, cxHoraInicio.Text, Formula );
        cdHorarios.Post;

        {$REGION 'Crear excel'}

        Cuadre.EncuentraHorario( cdHorarios, Fila );
        Ruta := Cuadre.Path + Cuadre.GenerarNombreAleatorio() + '.xls' ;

        try

          cxCuadre.SaveToFile( Ruta );

          Excel := CreateOleObject('Excel.Application');
          Excel.DisplayAlerts := False;
          Excel.ScreenUpdating := True;
          Excel.Visible := False;
          Excel.Workbooks.Open( Ruta );
          Libro := Excel.Workbooks[1];
          Hoja := Libro.Sheets[ 1 ] ;
          Hoja.Select;
          Hoja.Unprotect;

        except
          on e:Exception do
          begin
            MessageDlg('No se puede continuar verifique tener instalada la suite de Microsoft Office', mtInformation, [mbOK], 0);
            Exit;
          end;
        end;

        {$ENDREGION}

        Rango := Excel.Range[ 'C' + IntToStr( Fila ) ];
        Rango.Value := cxHoraInicio.Text;

        Rango := Excel.Range[ 'D' + IntToStr( Fila ) ];
        Rango.Value := cxHoraFinal.Text;

        if cxHoraFinal.Text = '24:00' then
          Rango.NumberFormat := '[h]:mm:ss'
        else
          Rango.NumberFormat := 'hh:mm';

        Rango := Excel.Range[ 'E' + IntToStr( Fila ) ];
        Rango.NumberFormat := '0.0000';
        Rango.Value := cdHorarios.FieldByName( 'Duracion' ).AsString;//Cuadre.ObtenerDuracion( cxHoraFinal.Text, cxHoraInicio.Text, Formula );

        Libro.ActiveSheet.Protect( True, True, True );
        Libro.SaveAs( Ruta, 56 );
        Excel.Quit;

        cxCuadre.LoadFromFile( Ruta );

        Cuadre.Cambios := True;

      end;

    end;
  finally
    
  end;  
end;

procedure TfrmCuadreCategoria.EliminarHorario(Sender: TObject);
var
  Fila : TExcelRow;
  Archivo : TFileName;
  No : Integer;
begin
  try
    Application.ProcessMessages;
    Fila := cxCuadre.ActiveCell.Y + 1;
    if ( Cuadre.EncuentraFolio( cdFolios, Fila ) ) and ( Cuadre.EncuentraActividad( cdActividades, Fila ) ) or ( Cuadre.EncuentraHorario( cdHorarios, Fila ) ) then
    begin

      if MessageDlg( '¿Desea realmente eliminar el horario?', mtConfirmation, [ mbYes, mbCancel ], 0 ) = mrYes then
      begin

        if Cuadre.EncuentraHorario( cdHorarios, Fila ) then
        begin
          cdHorarios.Delete;
          Cuadre.Cambios := True;
          Cuadre.ActualizarFilasConjuntoActual( cdFolios, cdActividades, -1 );
          Cuadre.ActualizarFilasFolios( cdFolios, -1 );
          Cuadre.ActualizarFilasActividades( cdActividades, -1 );
          Cuadre.ActualizarFilasHorarios( cdHorarios, -1, ftEliminar, Fila );
          Archivo := Cuadre.Path + Cuadre.GenerarNombreAleatorio() + '.xls';
          cxCuadre.SaveToFile( Archivo );
          cxCuadre.LoadFromFile( Cuadre.EliminarHorario( Fila, Archivo ) );
        end;

      end;

    end;
  finally
    
  end;  
end;

procedure TfrmCuadreCategoria.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  Cuadre.DesbloquearCategoria;
  Action := caFree;
end;

procedure TfrmCuadreCategoria.FormCreate(Sender: TObject);
begin

  Cuadre := TCuadre.Create;
  Cuadre.Tipo := 'Personal';
  MostrarCategorias( treeCategorias, ftPersonal, False );
  Cuadre.TipoCategoria := ftPersonal;

  zPernoctas.Active := False;
  zPernoctas.Open;
  zPlataformas.Active := False;
  zPlataformas.Open;
  Cuadre.DefinirEstructuras( cdFolios, cdActividades, cdCategorias, cdHorarios );

  Self.Caption := 'CUADRE - Contrato: ' + Cuadre.Contrato + ', OT: ' + Cuadre.Orden + ', Fecha: ' + Cuadre.Fecha + ' - ' + UpperCase( global_estado_reporte );

  btnGuardar.Visible := lowercase(global_estado_reporte) = 'pendiente';

end;

procedure TfrmCuadreCategoria.ImprimirEstructura1Click(Sender: TObject);
begin
  Cuadre.ExportarEstructura( cdFolios, cdActividades, cdHorarios );
end;

procedure TfrmCuadreCategoria.lstCuentasClick(Sender: TObject);
var
  ZAjuste : TZReadOnlyQuery;
begin
  InicializarZQuery( ZAjuste );
  ZAjuste.SQL.Text := 'select dCantidad from bitacoradepernocta '+
                     'where '+
                        'sContrato = :orden and '+
                        'dIdFecha = :fecha and '+
                        'sNumeroOrden = :folio and '+
                        'sIdCuenta = (select sIdCuenta from cuentas where sIdPernocta = :pernocta);';

  ZAjuste.ParamByName( 'orden' ).AsString := Cuadre.Orden;
  ZAjuste.ParamByName( 'fecha' ).AsString := Cuadre.Fecha;
  ZAjuste.ParamByName( 'folio' ).AsString := cdFolios.FieldByName( 'Folio' ).AsString;
  ZAjuste.ParamByName( 'pernocta' ).AsString := lstCuentas.Items[ lstCuentas.ItemIndex ];
  ZAjuste.Open;

  clcAjuste.Value := ZAjuste.FieldByName( 'dCantidad' ).AsFloat;
  btnAplicaAjuste.Enabled := False;
end;

procedure TfrmCuadreCategoria.ComprobarSuma(SumarValor: Double);
var
  Rango : TcxSSCellObject;

  Valor : Integer;
  Suma : Double;
begin
  try

    Application.ProcessMessages;
    Application.ProcessMessages;

    Cuadre.Cambios := True;

    Rango := cxCuadre.Sheet.GetCellObject( cxCuadre.Sheet.ActiveCell.X, cxCuadre.Sheet.ActiveCell.Y );
    if Length( VarToStr( Rango.CellValue ) ) > 0 then
    begin
      try
        Valor := StrToInt( VarToStr( Rango.CellValue ) );
      except
        raise Exception.Create( 'Valor Invalido' );
      end;

      if not Valor > 0 then
        raise Exception.Create( 'No se permiten valores negativos' );

      Rango := cxCuadre.Sheet.GetCellObject( 6, cdFolios.FieldByName( 'Inicio' ).AsInteger - 3 );
      Suma := StrToFloat( VarToStr( Rango.CellValue ) );

      Suma := Suma + SumarValor;

      lblSuma.Caption := FloatToStr( Suma );

      case CompareValue( Suma, cdCategorias.FieldByName( 'A_Bordo' ).AsFloat, 0.2 ) of
          LessThanValue    : Rango.Style.Brush.BackgroundColor := 41;
          EqualsValue      : Rango.Style.Brush.BackgroundColor := 42;
          GreaterThanValue : Rango.Style.Brush.BackgroundColor := 45;
      end;

    end;

  except
    on e:Exception do
      MessageDlg( e.Message, mtInformation, [ mbOK ], 0 );
  end;
end;

function TfrmCuadreCategoria.CategoriaBloqueada(Categoria: string ):Boolean;
var
  zBusca : TZReadOnlyQuery;
begin
  InicializarZQuery( zBusca );
  if Cuadre.TipoCategoria = ftPersonal then
    zBusca.SQL.Text := 'select eOcupado from personal where sContrato = :contrato and sIdPersonal = :categoria limit 1;'
  else
    zBusca.SQL.Text := 'select eOcupado from equipos where sContrato = :contrato and sIdEquipo = :categoria limit 1;';

  zBusca.ParamByName( 'contrato' ).AsString := Cuadre.Contrato;
  zBusca.ParamByName( 'categoria' ).AsString := Categoria;
  zBusca.Open;
  zBusca.First;

  Result := False;
  if zBusca.RecordCount = 1 then
    Result := LowerCase( zBusca.FieldByName( 'eOcupado' ).AsString ) = 'si';

end;

function TfrmCuadreCategoria.ObtenerFilaNuevoHorario():TExcelRow;
var
  IdRegistro : TIdRegistroHorario;
begin
  IdRegistro := cdHorarios.FieldByName( 'IdRegistro' ).AsInteger;
  cdHorarios.Filtered := False;
  cdHorarios.Filter := 'IdRegistroActividad = ' + cdActividades.FieldByName( 'IdRegistro' ).AsString ;
  cdHorarios.Filtered := True;
  cdHorarios.Last;

  if cdHorarios.RecordCount > 0 then
    Result := cdHorarios.FieldByName( 'Fila' ).AsInteger + 1
  else
    Result := cdActividades.FieldByName( 'Fin' ).AsInteger;

  cdHorarios.Filtered := False;
  cdHorarios.IndexFieldNames := EMPTY_STRING;
  cdHorarios.IndexFieldNames := 'Fila';

  cdHorarios.Locate( 'IdRegistro', IdRegistro, [] );

end;

end.
