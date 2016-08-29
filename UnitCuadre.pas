unit UnitCuadre;

interface

uses
  global, frm_connection,

  Windows, Messages, Dialogs, ZSqlProcessor, ZConnection,
  SysUtils, ComCtrls, Forms, cxGroupBox, cxEdit, Controls, ComObj, UnitExcel,
  cxTreeView, DB, ZAbstractRODataset, ZAbstractDataset, ZDataset, DBClient ;

type
  TftCategoria = ( ftPersonal = 0, ftEquipo = 1 );

type
  TftBloqueado = ( ftCategoriaBloqueada = 0, ftCategoriaLibre = 1 );

type
  TftFormula = ( ftA_Bordo = 0, ftTierra = 1, ftTiempoExtra = 2, ftEq = 3, ftOtro = 4 );

type
  TftTipoOperacion = ( ftInsert = 0, ftEliminar = 1 );

type
  TExcelColIndex = type Integer;
  TExcelRow = type Integer;
  TExcelColAlias = type string;
  TExcelRangeAlias = type string;
  TExcelFormula = type string;
  TExcelInstance = type Variant;

  TIdRegistroFolio = type Integer;
  TIdRegistroActividad = type Integer;
  TIdRegistroCategoria = type Integer;
  TIdRegistroHorario = type Integer;


//Clase personalizada para columnas en excel
type
  TExcelColumna = class( TObject )
    Column : TExcelColIndex;
    StrColumn : TExcelColAlias;
    function Columna():TExcelColAlias;
    function _Columna():TExcelColAlias;overload;
    function _Columna( Increment : TExcelColIndex ):TExcelColAlias;overload;
    function Columna_():TExcelColAlias;overload;
    function Columna_( Increment : TExcelColIndex ):TExcelColAlias;overload;

    constructor Create;
  end;

//Clase personalizada para fila en excel
type
  TExcelFila = class( TObject )
    public
      Row,
      OldRow : TExcelRow;

    function SigFila( count : TExcelRow = 1 ):TExcelRow;
    function StrFila:string;overload;
    function StrFila( Increment : Integer ):string;overload;

  end;

//Clase para datos y metodos del Cuadre
type
  TCuadre = class( TObject )
    public
      Categoria,
      Contrato,
      Orden,
      Fecha,
      Tipo,
      Path : string;

      IdCategoria : TIdRegistroCategoria;

      Formula : TftFormula;
      TipoCategoria : TftCategoria;

      Cambios : Boolean;

      constructor Create;
      procedure CargarCategorias( var cdCategoria : TClientDataSet; TipoRecurso : TftCategoria );
      procedure CargarFolios( var cdFolio : TClientDataSet );
      procedure CargarActividades( var cdActividad, cdFolios : TClientDataSet; Formula : TftFormula );
      procedure CargarHorarios( var cdFolio, cdActividad, cdHorarios, cdCategoria : TClientDataSet );
      procedure DefinirEstructuras( var cdFolio, cdActividad, cdCategoria, cdHorario : TClientDataSet );
      procedure ActualizarFilasHorarios( var cdHorarios : TClientDataSet; Incremento : Integer = 1 ; Operacion : TftTipoOperacion = ftInsert; Fila : Integer = -1 );
      procedure ActualizarFilasActividades( var cdActividades : TClientDataSet; Incremento : Integer = 1 );
      procedure ActualizarFilasFolios( var cdFolios : TClientDataSet; Incremento : Integer = 1 );
      procedure ActualizarFilasConjuntoActual( var cdFolio, cdActividad : TClientDataSet; Incremento : Integer = 1 );
      procedure ExportarEstructura( Folios, Actividades, Horarios : TClientDataSet );
      procedure BloquearCategoria;
      procedure DesbloquearCategoria;

      function GenerarCuadre( var cdFolios, cdCategorias, cdActividades, cdHorarios : TClientDataSet; var prgF, prgA, prgH : TProgressBar; Formulas : TftFormula ):TFileName;
      function EncuentraFolio( var cdFolios : TClientDataSet; Fila : TExcelRow ):Boolean ;
      function EncuentraHorario( var cdHorarios : TClientDataSet; Fila : TExcelRow ):Boolean;
      function EncuentraActividad( var cdActividades : TClientDataSet; Fila : TExcelRow ):Boolean;
      function CrearHorario( IdActividad : TIdRegistroActividad; var cdHorario : TClientDataSet; HInicio, HFin : string; Fila : TExcelRow; Formulas : TftFormula ):TIdRegistroHorario;
      function GenerarNombreAleatorio( LongitudCaracteres : Integer = 15 ):TFileName;
      function RegenerarCuadre( var cdFolios, cdActividades, cdHorario : TClientDataSet; DirArchivo : TFileName; Row : Integer; HInicio, HFin, Duracion : string; Formulas : TftFormula ):TFileName;
      function RegenerarFormulaSuma( var ExcelApp : TExcelInstance; cdFolios, cdActividades : TClientDataSet ):TExcelFormula;
      function GuardarCuadre( var cdFolios, cdActividades, cdCategoria, cdHorarios : TClientDataSet; Ruta : TFileName ):TFileName;
      function EliminarHorario( Fila : TExcelRow; Archivo : TFileName ):TFileName;
      function ObtenerDuracion( HFin, HInicio : String; Formulas : TftFormula ):string;

    private
      FZGuardaCuadre : TZSQLProcessor;

  end;

  {$region 'Procedimientos'}

  procedure InicializarZQuery( var zQuery : TZQuery );overload;
  procedure InicializarZQuery( var zQuery : TZReadOnlyQuery );overload;
  procedure MostrarCategorias( var TreeRecursos : TcxTreeView; TipoRecurso : TftCategoria; ConservarDatos : Boolean = True );
  procedure CargarCuadresHechos( var TreeDias : TcxTreeView );
  procedure InicializarForm( var Ventana : TForm; Grupo : TcxGroupBox; Largo, Ancho : Integer );
  procedure ReestablecerBox( var Box : TcxGroupBox; NuevoParent : TForm );

  {$endregion}

  {$region 'Constantes'}

const
  TIPO_CATEGORIA : array[ 0..1 ] of string = ( 'Personal', 'Equipo' );
  EMPTY_STRING : string = '';
  TABLA_RECURSO : array[ 0..1 ] of string = ( 'bitacoradepersonal', 'bitacoradeequipos' );
  TODAS_LAS_CATEGORIAS : Integer = 0;
  UNA_CATEGORIA : integer = 1;
  TODOS_LOS_FOLIOS : Integer = 1;
  TODAS_LAS_ACTIVIDADES :  Integer = 0;
  FILA_NO_ASIGNADA : Integer = -1;
  DURACION_NO_ASIGNADA : Double = 0.0;
  ALFABETO = 'ABCDEFGHJKLMNPQRSTUVWXYZ';

  {$endregion}

  {$region 'Consultas SQL'}

  {$region 'Consulta Reportes Diarios con Actividades'}

  SQL_REPORTES : string = 'select  rd.dIdFecha as Fecha '+
                          'from reportediario rd '+
                          'inner join turnos t '+
                            'on ( t.sContrato = rd.sOrden '+
                                'and t.sIdTurno = rd.sIdTurno ) '+
                          'inner join bitacoradeactividades ba '+
                            'on ( ba.sContrato = rd.sOrden '+
                                'and ba.dIdFecha = rd.dIdFecha ) '+
                          'inner join actividadesxorden ao '+
                            'on ( ao.sContrato = ba.sContrato '+
                                'and ao.sWbs = ba.sWbs '+
                                'and ao.sNumeroActividad = ba.sNumeroActividad '+
                                'and ao.sTipoActividad = "Actividad" ) '+
                          'inner join ordenesdetrabajo ot '+
                            'on ( ot.sContrato = ba.sContrato '+
                                'and ot.sNumeroOrden = ba.sNumeroOrden ) '+
                          'where rd.sOrden = :orden '+
                            'and ba.sIdTipoMovimiento = "ED" '+
                          'group by rd.dIdFecha '+
                          'order by rd.dIdFecha desc ';

  {$endregion}

  {$region 'Consultas MOE'}

  SQL_CATEGORIAS : array[ 0..1 ] of string = ('select mr.iIdMoe, '+
                                              'mr.sIdRecurso, '+
                                              'mr.sDescripcion, '+
                                              'ax.sTierra, '+
                                              'mr.dCantidad as iSolicitado, '+
                                              'p.sIdTipoPersonal as sAgrupa, '+
                                              'p.eOcupado, '+
                                              'mra.dCantidad as iAbordo, p.sIdTipoPersonal, p.iItemOrden, '+
                                              'if ( lower( mr.sDescripcion ) = "tiempo extra", "Si" , "No") as sTE '+

                                            'from moerecursos mr '+

                                            'inner join moe m '+
                                             'on ( m.sContrato = :orden and '+
                                                  'mr.iIdMoe = m.iIdMoe ) '+

                                            'inner join moerecursos_abordo mra '+
                                             'on ( mra.iIdMoe = m.iIdMoe and '+
                                                  'mra.sIdRecurso = mr.sIdRecurso and '+
                                                  'mra.eTipoRecurso = "Personal" ) '+

                                            'inner join personal p '+
                                             'on ( p.sContrato = :contrato and '+
                                                  'p.sIdPersonal = mr.sIdRecurso and '+
                                                  'p.lCobro = "Si" ) '+
                                            '       '+
                                            'left join tiposdepersonal as tp '+
                                              'on ( tp.sIdTipoPersonal = p.sIdTipoPersonal ) '+

                                            'inner join anexos ax '+
                                              'on ( ax.sAnexo = p.sAnexo && '+
                                                   'lower( ax.sTipo ) = "personal" ) '+

                                            'where '+
                                              'm.dIdFecha = ( select max( m1.dIdFecha ) '+
                                                             'from moe m1 '+
                                                             'where m1.sContrato = :orden '+
                                                             'and m1.dIdFecha <= :fecha ) and '+
                                              'mr.eTipoRecurso = "Personal" and '+
                                              'if( :todas_categorias = 1, mr.sIdRecurso = :id_recurso, mr.sIdRecurso <> "!" ) '+

                                            'group by p.sIdPersonal; '

                                            ,

                                              'select '+
                                              'mr.iIdMoe, '+
                                              'mr.sIdRecurso, '+
                                              'mr.sDescripcion, '+
                                              'e.sIdTipoEquipo as sAgrupa, '+
                                              'e.eOcupado, '+
                                              'mr.dCantidad as iSolicitado, '+
                                              'mra.dCantidad as iAbordo, e.iItemOrden, '+
                                              '"No" as sTE, '+
                                              '"No" as sTierra '+

                                            'from moerecursos mr '+

                                            'inner join moe m '+
                                              'on ( m.sContrato = :orden and '+
                                                   'mr.iIdMoe = m.iIdMoe ) '+

                                            'inner join moerecursos_abordo mra '+
                                              'on ( mra.iIdMoe = m.iIdMoe and '+
                                                   'mra.sIdRecurso = mr.sIdRecurso and '+
                                                   'mra.eTipoRecurso = "Equipo" ) '+

                                            'inner join equipos e '+
                                              'on ( e.sContrato = :contrato and '+
                                                   'mr.eTipoRecurso = "Equipo" and '+
                                                   'e.lCobro = "Si" and '+
                                                   'e.sIdEquipo = mr.sIdRecurso ) '+

                                            'left join tiposdeequipo as te '+
                                              'on ( te.sIdTipoEquipo = e.sIdTipoEquipo ) '+

                                            'where '+
                                              'm.dIdFecha = ( select max( m1.dIdFecha ) '+
                                                             'from moe m1 '+
                                                             'where m1.sContrato = :orden '+
                                                             'and m1.dIdFecha <= :fecha ) and '+
                                              'if( :todas_categorias = 1, mr.sIdRecurso = :id_recurso, mr.sIdRecurso <> "!" );; ' );

  {$endregion}

  {$region 'Consulta Cuadre existente'}

  SQL_CUADRES_EXISTENTES : string = 'select '+
                                      'distinct rd.dIdFecha '+
                                    'from reportediario as rd '+

                                    'inner join bitacoradeactividades as ba '+
                                      'on( ba.sContrato = rd.sOrden and '+
                                          'ba.dIdFecha = rd.dIdFecha and '+
                                          'ba.sIdTipoMovimiento = "ED" ) '+

                                    'inner join bitacoradepersonal as bp '+
                                      'on ( ba.sContrato = bp.sContrato and '+
                                           'ba.dIdFecha = bp.dIdFecha and '+
                                           'ba.sNumeroActividad = bp.sNumeroActividad and '+
                                           'ba.iIdActividad = bp.iIdActividad ) '+

                                    'inner join personal as p '+
                                      'on ( p.sContrato = :contrato and '+
                                           'p.sIdPersonal = bp.sIdPersonal and '+
                                           'p.lPernocta = "Si" ) '+

                                    'inner join tiposdepersonal as tp '+
                                      'on ( tp.sIdTipoPersonal = p.sIdTipoPersonal ) '+

                                    'where '+
                                      'rd.sContrato = :contrato and '+
                                      'rd.sOrden = :orden '+

                                    'order by rd.dIdFecha desc ;';

  {$endregion}

  {$region 'Consulta Folios reportados'}

  SQL_FOLIOS_REPORTADOS : string = 'select ot.sNumeroOrden as sIdFolio, '+
                                           'ot.sNumeroOrden, '+
                                           'ot.iJornadas '+
                                    'from bitacoradeactividades ba '+
                                    'inner join reportediario rd '+
                                      'on ( rd.sOrden = ba.sContrato '+
                                        'and rd.dIdFecha = ba.dIdFecha ) '+
                                    'inner join ordenesdetrabajo ot '+
                                      'on ( ot.sContrato = ba.sContrato '+
                                        'and ot.sNumeroOrden = ba.sNumeroOrden ) '+
                                    'inner join contratos c '+
                                      'on ( c.sContrato = ba.sContrato ) '+
                                    'where ba.sContrato = :orden '+
                                    'and ba.sIdTipoMovimiento = "ED" '+
                                    'and ba.dIdFecha = :fecha '+
                                    'group by ot.sNumeroOrden ';

  {$endregion}

  {$region 'Consulta Actividades Y Folios Reportados'}

  SQL_ACTIVIDADES_Y_FOLIOS_REPORTADOS : string = 'select '+
                                                 'ba.sContrato, '+
                                                 'ba.sNumeroOrden, '+
                                                 'ba.iIdDiario, '+
                                                 'ba.iIdActividad, '+
                                                 'ba.sNumeroActividad, '+
                                                 'ba.sHoraInicio, '+
                                                 'ba.sHoraFinal, '+
                                                 'time_format( timediff( ba.sHoraFinal, ba.sHoraInicio ), '+
                                                              'if( :formula = 2, "%H.%i", "%H:%i" ) ) as Duracion, '+
                                                 'ba.sWbs, '+
                                                 'ba.iIdTarea, '+
                                                 'ba.mDescripcion, '+
                                                 'fl.sIdPernocta, '+
                                                 'fl.sIdPlataforma, '+
                                                 '( select ifnull (sum( dCantidad ), 0 ) as dCantidad '+
                                                   'from bitacoradepernocta '+
                                                   'where sContrato= ba.sContrato and '+
                                                   'dIdFecha = ba.dIdFecha and '+
                                                   'sNumeroOrden = ba.sNumeroOrden limit 1 ) as Ajuste '+
 
                                               'from bitacoradeactividades as ba '+
 
                                               'inner join reportediario as rd '+
                                                 'on ( rd.sOrden = ba.sContrato and '+
                                                      'rd.dIdFecha = ba.dIdFecha ) '+
 
                                               'inner join ordenesdetrabajo as fl '+
                                                 'on ( fl.sContrato = ba.sContrato and '+
                                                      'fl.sNumeroOrden = ba.sNumeroOrden ) '+
                                               'where '+
                                                 'ba.sContrato = :orden and '+
                                                 'ba.dIdFecha = :fecha and '+
                                                 'ba.sIdTipoMovimiento = "ED" '+
 
                                               'group by '+
                                                 'if ( :x_folio = 1, ba.sNumeroOrden, ba.iIdActividad ) '+
 
                                               'order by '+
                                                 'ba.sNumeroOrden, '+
                                                 'ba.sNumeroActividad, '+
                                                 'ba.sHoraInicio; ';

  {$endregion}

  {$region 'Consulta Inserta en bitacoras'}

  SQL_INSERT : array[0..1] of string = ('insert into bitacoradepersonal '+
                                          '(sContrato, '+
                                          'dIdFecha, '+
                                          'iIdDiario, '+
                                          'iItemOrden, '+
                                          'sIdPersonal, '+
                                          'sTipoObra, '+
                                          'sDescripcion, '+
                                          'sIdPernocta, '+
                                          'sIdPlataforma, '+
                                          'sHoraInicio, '+
                                          'sHoraFinal, '+
                                          'dCantidad, '+
                                          'sTipoPernocta, '+
                                          'sWbs, '+
                                          'sNumeroActividad, '+
                                          'dCantHH, '+
                                          'sNumeroOrden, '+
                                          'iIdActividad, '+
                                          'iIdTarea, '+
                                          'sAgrupaPersonal, '+
                                          'sHoraInicioG, '+
                                          'sHoraFinalG ) '+#10+

                                    'values (:orden, '+
                                        ':fecha, '+
                                        ':iddiario, '+
                                        ':ItemOrden, '+
                                        ':idrecurso, '+
                                        '"PU", '+
                                        ':descripcion, '+
                                        ':pernocta, '+
                                        ':plataforma, '+
                                        ':hinicio, '+
                                        ':hfinal, '+
                                        ':cantidad, '+
                                        '"6.1.", '+
                                        ':wbs, '+
                                        ':actividad, '+
                                        ':cantidadhh, '+
                                        ':folio, '+
                                        ':idactividad, '+
                                        ':tarea, '+
                                        ':Categoria, '+
                                        ':hinicio, '+
                                        ':hfinal ) '

                                        ,

                                        'insert into bitacoradeequipos (sContrato, '+
                                        'dIdFecha, '+
                                        'iIdDiario, '+
                                        'iItemOrden, '+
                                        'sIdEquipo, '+
                                        'sDescripcion, '+
                                        'sIdPernocta, '+
                                        'sTipoObra, '+
                                        'sHoraInicio, '+
                                        'sHoraFinal, '+
                                        'dCantidad, '+
                                        'sWbs, '+
                                        'sNumeroActividad, '+
                                        'dCantHH, '+
                                        'sNumeroOrden, '+
                                        'iIdActividad, '+
                                        'iIdTarea, '+
                                        'sHoraInicioG, '+
                                        'sHoraFinalG ) '+#10+

                                    'values (:orden, '+
                                        ':fecha, '+
                                        ':iddiario, '+
                                        ':ItemOrden, '+
                                        ':idrecurso, '+
                                        ':descripcion, '+
                                        ':pernocta, '+
                                        '"PU", '+
                                        ':hinicio, '+
                                        ':hfinal, '+
                                        ':cantidad, '+
                                        ':wbs, '+
                                        ':actividad, '+
                                        ':cantidadhh, '+
                                        ':folio, '+
                                        ':idactividad, '+
                                        ':tarea, '+
                                        ':hinicio, '+
                                        ':hfinal )');

  {$endregion}

  {$region 'Consulta Borra de las bitacoras'}

  SQL_LIMPIA_BITACORA : array[ 0..1 ] of string = ( 'delete from bitacoradepersonal where sContrato = :orden and dIdFecha = :fecha and sIdPersonal = :categoria; ',
                                                     'delete from bitacoradeequipos where sContrato = :orden and dIdFecha = :fecha and sIdEquipo = :categoria; ' );

  {$endregion}

  {$region 'Consulta Personal y Equipo Existente'}

  SQL_CUADRE_EXISTENTE : array[0..1] of string = ('select ot.sNumeroOrden, '+
                               'ba.iIdActividad, '+
                               'bp.sHoraInicio, '+
                               'bp.sHoraFinal, '+
                               'time_format( timediff( bp.sHoraFinal, bp.sHoraInicio ), '+
                                            'if( replace( upper( p.sDescripcion ), " ", "" ) = "TIEMPOEXTRA", '+
                                            '"%H.%i", "%H:%i" ) ) as Duracion,'+
                               'bp.iIdDiario, '+
                               'ba.sNumeroActividad, '+
                               'ba.mDescripcion, '+
                               'p.sIdPersonal as sIdRecurso, '+
                               'p.sDescripcion, '+
                               'bp.dCantidad, '+
                               'bp.dCantHH, '+
                               'bp.sIdPernocta, '+
                               'bp.sIdPlataforma '+ #10 +

                        'from bitacoradepersonal bp '+
                        'inner join bitacoradeactividades ba '+
                          'on ( bp.sContrato = ba.sContrato '+
                            'and bp.iIdActividad = ba.iIdActividad '+
                            'and bp.sNumeroActividad = ba.sNumeroActividad '+
                            'and bp.iIdDiario = ba.iIdDiario '+
                            'and bp.sNumeroOrden = ba.sNumeroOrden '+
                            'and ba.sIdTipoMovimiento = "ED" '+

                          ') '+ #10 +

                        'inner join personal p '+
                          'on ( p.sContrato = :contrato '+
                            'and p.sIdPersonal = bp.sIdPersonal '+
                            'and p.lCobro = "Si" '+
                          ') '+ #10 +

                        'inner join ordenesdetrabajo ot '+
                          'on ( ot.sNumeroOrden = ba.sNumeroOrden ) '+ #10 +

                        'inner join plataformas pl '+
                          'on ( ot.sIdPlataforma = pl.sIdPlataforma ) '+ #10 +

                        'inner join pernoctan pr '+
                          'on ( ot.sIdPernocta = pr.sIdPernocta ) '+ #10 +

                        'where bp.sContrato = :orden '+
                        'and bp.dIdFecha = :fecha '+
                        'and bp.sIdPersonal = :idrecurso '+

                        'order by ba.iIdActividad, p.sIdPersonal '

                        ,

                        'select ot.sNumeroOrden, '+
                                 'ba.iIdActividad, '+
                                 'be.sHoraInicio, '+
                                 'be.sHoraFinal, '+
                                 'time_format( timediff( be.sHoraFinal, be.sHoraInicio ), "%H:%i" ) as Duracion, '+
                                 'ba.iIdDiario, '+
                                 'ba.sNumeroActividad, '+
                                 'ba.mDescripcion, '+
                                 'e.sIdEquipo  as sIdRecurso, '+
                                 'e.sDescripcion, '+
                                 'be.dCantidad, '+
                                 'be.dCantHH, '+
                                 'be.sIdPernocta '+ #10 +

                          'from bitacoradeequipos be '+
                          'inner join bitacoradeactividades ba '+
                            'on ( be.sContrato = ba.sContrato '+
                              'and be.iIdActividad = ba.iIdActividad '+
                              'and be.sNumeroActividad = ba.sNumeroActividad '+
                              'and be.iIdDiario = ba.iIdDiario '+
                              'and be.sNumeroOrden = ba.sNumeroOrden '+
                              'and ba.sIdTipoMovimiento = "ED" '+
                            ') '+ #10 +

                          'inner join equipos e '+
                            'on ( e.sContrato = :contrato '+
                              'and e.sIdEquipo = be.sIdEquipo '+
                              'and e.lCobro = "Si" '+
                            ') '+ #10 +

                          'inner join ordenesdetrabajo ot '+
                            'on ( ot.sNumeroOrden = ba.sNumeroOrden ) '+ #10 +

                          'inner join plataformas pl '+
                            'on ( ot.sIdPlataforma = pl.sIdPlataforma ) '+ #10 +

                          'inner join pernoctan pr '+
                            'on ( ot.sIdPernocta = pr.sIdPernocta ) '+ #10 +

                          'where be.sContrato = :orden '+
                          'and be.dIdFecha = :fecha '+
                          'and be.sIdEquipo = :idrecurso ' );

  {$endregion}


  {$endregion}

var
  Cuadre : TCuadre;

implementation

  constructor TCuadre.Create;
  begin
    inherited Create;
    Contrato := global_contrato_barco;
    Orden := global_contrato;
    Fecha := FormatDateTime( 'YYYY-MM-DD', global_fecha );
    FZGuardaCuadre := TZSQLProcessor.Create( nil );
    FZGuardaCuadre.Connection := connection.zConnection;
    FZGuardaCuadre.Script.Clear;

  end;

  function TCuadre.EncuentraFolio(var cdFolios: TClientDataSet; Fila: TExcelRow):Boolean;
  var
    IdRespaldo : TIdRegistroFolio;
  begin
    IdRespaldo := cdFolios.FieldByName( 'IdRegistro' ).AsInteger;
    cdFolios.First;

    if not cdFolios.Locate( 'Fila', Fila, [] ) then
    begin

      Result := False;
      while not cdFolios.Eof do
      begin
        if ( Fila >= cdFolios.FieldByName( 'Inicio' ).AsInteger ) and ( Fila <= cdFolios.FieldByName( 'Fin' ).AsInteger ) then
        begin
          Result := True;
          Break;
        end;
        cdFolios.Next;
      end;

    end;

    if not Result then
      cdFolios.Locate( 'IdRegistro', IdRespaldo, [] );

  end;

  function TCuadre.EncuentraHorario(var cdHorarios: TClientDataSet; Fila: TExcelRow):Boolean;
  begin
    Result := cdHorarios.Locate( 'Fila', Fila, [] );
  end;

  function TCuadre.EncuentraActividad(var cdActividades: TClientDataSet; Fila: TExcelRow):Boolean;
  var
    IndexActividad : Integer;
    IdRespaldo : TIdRegistroActividad;
  begin

    IdRespaldo := cdActividades.FieldByName( 'IdRegistro' ).AsInteger;
    cdActividades.First;
    Result := True;

    if not cdActividades.Locate( 'Fila', Fila, [] ) then
    begin

      Result := False;
      while not cdActividades.Eof do
      begin
        if ( Fila >= cdActividades.FieldByName( 'Inicio' ).AsInteger ) and ( Fila <= cdActividades.FieldByName( 'Fin' ).AsInteger ) then
        begin
          Result := True;
          Break;
        end;
        cdActividades.Next;

      end;

    end;

    if not Result then
      cdActividades.Locate( 'IdRegistro', IdRespaldo, [] );  

  end;

  function TCuadre.CrearHorario( IdActividad : TIdRegistroActividad; var cdHorario : TClientDataSet; HInicio, HFin : string; Fila : TExcelRow; Formulas : TftFormula ):TIdRegistroHorario;
  begin

    cdHorario.IndexFieldNames := 'Fila';
    cdHorario.Append;
    cdHorario.FieldByName( 'IdRegistro' ).AsInteger := cdHorario.RecordCount + 1;
    Result := cdHorario.FieldByName( 'IdRegistro' ).AsInteger;
    cdHorario.FieldByName( 'IdRegistroActividad' ).AsInteger := IdActividad;
    cdHorario.FieldByName( 'IdRegistroCategoria' ).AsInteger := Cuadre.IdCategoria;
    cdHorario.FieldByName( 'Inicio' ).AsString := HInicio;
    cdHorario.FieldByName( 'Fin' ).AsString := HFin;
    cdHorario.FieldByName( 'Duracion' ).AsString := ObtenerDuracion( HFin, HInicio, Formulas );
    cdHorario.FieldByName( 'Cantidad' ).AsInteger := 0;
    cdHorario.FieldByName( 'CantidadHorasHombre' ).AsInteger := 0;
    cdHorario.FieldByName( 'Fila' ).AsInteger := Fila;
    cdHorario.Post;
    
    cdHorario.IndexFieldNames := 'Fila';
    Application.ProcessMessages;

    cdHorario.First;
    while not cdHorario.Eof do
    begin
      if cdHorario.FieldByName( 'Fila' ).AsInteger = Fila then
        Break;

      cdHorario.Next;
    end;

  end;

  function TCuadre.GenerarNombreAleatorio( LongitudCaracteres : Integer = 15 ):TFileName;
  var
    C : Char;
    Nombre : TFileName;
    Count : Integer;
  begin
    Nombre := EMPTY_STRING;
    for Count := 1 to LongitudCaracteres - 1 do
      Nombre := Nombre + Trim( ALFABETO[ Random( Length( ALFABETO ) ) + 1 ] );

    Result := Nombre;

  end;

  procedure TCuadre.CargarFolios( var cdFolio : TClientDataSet );
  var
    zFolios : TZReadOnlyQuery;
  begin
    try
      InicializarZQuery( zFolios );
      zFolios.SQL.Text := SQL_ACTIVIDADES_Y_FOLIOS_REPORTADOS;
      zFolios.ParamByName( 'orden' ).AsString := Cuadre.Orden;
      zFolios.ParamByName( 'fecha' ).AsString := Cuadre.Fecha;
      zFolios.ParamByName( 'x_folio' ).AsInteger := TODOS_LOS_FOLIOS;
      zFolios.Open;
      zFolios.First;

      if cdFolio.RecordCount > 0 then
        cdFolio.EmptyDataSet;

      while not zFolios.Eof do
      begin
        cdFolio.Append;
        cdFolio.FieldByName( 'IdRegistro' ).AsInteger := cdFolio.RecordCount + 1;
        cdFolio.FieldByName( 'Folio' ).AsString := zFolios.FieldByName( 'sNumeroOrden' ).AsString;
        cdFolio.FieldByName( 'Fila' ).AsInteger := FILA_NO_ASIGNADA;
        cdFolio.FieldByName( 'Inicio' ).AsInteger := FILA_NO_ASIGNADA;
        cdFolio.FieldByName( 'Fin' ).AsInteger := FILA_NO_ASIGNADA;
        cdFolio.FieldByName( 'Pernocta' ).AsString := zFolios.FieldByName( 'sIdPernocta' ).AsString;
        cdFolio.FieldByName( 'Plataforma' ).AsString := zFolios.FieldByName( 'sIdPlataforma' ).AsString;
        cdFolio.FieldByName( 'Ajuste' ).AsFloat := StrToFloat( zFolios.FieldByName( 'Ajuste' ).AsString );
        cdFolio.Post;

        zFolios.Next;
      end;

    finally
      zFolios.Free;
    end;


  end;

  procedure TCuadre.ActualizarFilasHorarios( var cdHorarios : TClientDataSet; Incremento : Integer = 1 ; Operacion : TftTipoOperacion = ftInsert; Fila : Integer = -1 );
  begin

    cdHorarios.First;
    while not cdHorarios.Eof do
    begin

      if cdHorarios.FieldByName( 'Fila' ).AsInteger > Fila then
      begin
        cdHorarios.Edit;
        cdHorarios.FieldByName( 'Fila' ).AsInteger := cdHorarios.FieldByName( 'Fila' ).AsInteger + Incremento;
        cdHorarios.Post;
      end;

      cdHorarios.Next;

    end;

  end;

  procedure TCuadre.ActualizarFilasActividades(var cdActividades: TClientDataSet; Incremento: Integer = 1);
  var
    IdRegistro : TIdRegistroActividad;
  begin
    IdRegistro := cdActividades.FieldByName( 'IdRegistro' ).AsInteger;
    cdActividades.Next;

    while not cdActividades.Eof do
    begin
      cdActividades.Edit;
      cdActividades.FieldByName( 'Fila' ).AsInteger := cdActividades.FieldByName( 'Fila' ).AsInteger + Incremento;
      cdActividades.FieldByName( 'Inicio' ).AsInteger := cdActividades.FieldByName( 'Inicio' ).AsInteger + Incremento;
      cdActividades.FieldByName( 'Fin' ).AsInteger := cdActividades.FieldByName( 'Fin' ).AsInteger + Incremento;
      cdActividades.Post;
      cdActividades.Next;
    end;

    cdActividades.Locate( 'IdRegistro', IdRegistro, [] );

  end;

  procedure TCuadre.ActualizarFilasFolios(var cdFolios: TClientDataSet; Incremento: Integer = 1);
  var
    IdRegistro : TIdRegistroActividad;
  begin
    IdRegistro := cdFolios.FieldByName( 'IdRegistro' ).AsInteger;
    cdFolios.Next;

    while not cdFolios.Eof do
    begin
      cdFolios.Edit;
      cdFolios.FieldByName( 'Fila' ).AsInteger := cdFolios.FieldByName( 'Fila' ).AsInteger + Incremento;
      cdFolios.FieldByName( 'Inicio' ).AsInteger := cdFolios.FieldByName( 'Inicio' ).AsInteger + Incremento;
      cdFolios.FieldByName( 'Fin' ).AsInteger := cdFolios.FieldByName( 'Fin' ).AsInteger + Incremento;
      cdFolios.Post;
      cdFolios.Next;
    end;

    cdFolios.Locate( 'IdRegistro', IdRegistro, [] );

  end;

  procedure TCuadre.ActualizarFilasConjuntoActual(var cdFolio: TClientDataSet; var cdActividad: TClientDataSet; Incremento: Integer = 1 );
  begin

    cdFolio.Edit;
    cdFolio.FieldByName( 'Fin' ).AsInteger := cdFolio.FieldByName( 'Fin' ).AsInteger + Incremento;
    cdFolio.Post;

    cdActividad.Edit;
    cdActividad.FieldByName( 'Fin' ).AsInteger := cdActividad.FieldByName( 'Fin' ).AsInteger + Incremento;
    cdActividad.Post;

  end;

  procedure TCuadre.ExportarEstructura( Folios, Actividades, Horarios : TClientDataSet );
  var
    Excel,
    Libro,
    Hoja,
    Rango : TExcelInstance;

    Fila : TExcelFila;
    Columna : TExcelColumna;

    Archivo : TFileName;

    IdFolio : TIdRegistroFolio;
    IdActividad : TIdRegistroActividad;
    IdHorario : TIdRegistroHorario;

  begin

    try

      {$region 'Crear Objetos'}

      {$REGION 'Crear excel'}

      try
        Excel := CreateOleObject('Excel.Application');
        Libro := Excel.Workbooks.Add;

        Libro.Sheets.Add;
        while Libro.Sheets.Count > 1 do
          Libro.Sheets[1].Delete;

        Hoja := Libro.Sheets[1];
        Excel.DisplayAlerts := False;
        Excel.ScreenUpdating := True;
        Excel.Visible :=  True;
        Excel.Workbooks[1].Sheets[1].Name := 'CUADRE';

      except
        on e:Exception do
        begin
          MessageDlg('No se puede continuar verifique tener instalada la suite de Microsoft Office', mtInformation, [mbOK], 0);
          Exit;
        end;
      end;

      {$ENDREGION}

      Fila := TExcelFila.Create;
      Columna := TExcelColumna.Create;

      {$endregion}

      IdFolio := Folios.FieldByName( 'idregistro' ).asinteger;
      IdActividad := Actividades.FieldByName( 'idregistro' ).AsInteger;
      IdHorario := Horarios.FieldByName( 'idregistro' ).AsInteger;
      Horarios.IndexFieldNames := EmptyStr;

      Folios.First;
      while not Folios.Eof do
      begin

        Fila.Row := Folios.FieldByName( 'Fila' ).AsInteger;

        Excel.Range[ 'A' + Fila.StrFila ].Value := 'FOLIO';
        Excel.Range[ 'B' + Fila.StrFila ].Value := Folios.FieldByName( 'Folio' ).AsString;
        Excel.Range[ 'A' + Folios.FieldByName( 'Inicio' ).AsString + ':B' + Folios.FieldByName( 'Fin' ).AsString ].Interior.ColorIndex := 45;

        Folios.Next;
      end;

      Actividades.First;
      while not Actividades.Eof do
      begin
        Fila.Row := Actividades.FieldByName( 'Fila' ).AsInteger;

        Excel.Range[ 'C' + Fila.StrFila ].Value := 'ACTIVIDAD';
        Excel.Range[ 'D' + Fila.StrFila ].Value := Actividades.FieldByName( 'Actividad' ).AsString;
        Excel.Range[ 'C' + Actividades.FieldByName( 'Inicio' ).AsString + ':D' + Actividades.FieldByName( 'Fin' ).AsString ].Interior.ColorIndex := 46;

        Actividades.Next;

      end;

      Horarios.First;
      while not Horarios.Eof do
      begin
        Fila.Row := Horarios.FieldByName( 'Fila' ).AsInteger;

        Excel.Range[ 'E' + Fila.StrFila ].Value := 'HORARIO';
        Excel.Range[ 'F' + Fila.StrFila ] := Horarios.FieldByName( 'Inicio' ).AsString;
        Excel.Range[ 'G' + Fila.StrFila ] := Horarios.FieldByName( 'Fin' ).AsString;

        Horarios.Next;

      end;

      Folios.Locate( 'IdRegistro', IdFolio, [] );
      Actividades.Locate( 'IdRegistro', IdActividad, [] );
      Horarios.Locate( 'IdRegistro', IdHorario, [] );

    finally

    end;
  end;

  function TCuadre.EliminarHorario( Fila: TExcelRow; Archivo : TFileName ):TFileName;
  var
    Excel,
    Libro,
    Hoja : TExcelInstance;
  begin

    try
      {$region 'Crear excel'}

      try
        Excel := CreateOleObject('Excel.Application');
        Excel.DisplayAlerts := False;
        Excel.ScreenUpdating := True;
        Excel.Visible := False;
        Excel.Workbooks.Open( Archivo );
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

      {$endregion}

      
      Excel.Rows[ Fila ].Delete;

    finally
      Libro.ActiveSheet.Protect( True, True, True );
      Libro.SaveAs( Archivo, 56 );
      Application.ProcessMessages;
      Excel.Quit;
      Result := Archivo;
    end;

  end;


  procedure TCuadre.BloquearCategoria;
  begin
    if Length( Trim( Cuadre.Categoria ) ) > 0 then
    begin
      try
        connection.zCommand.Active := False;

        if Cuadre.TipoCategoria = ftPersonal then
          connection.zCommand.SQL.Text := 'update personal set eOcupado = "Si" where sContrato = :contrato and sIdPersonal = :categoria;'
        else
          connection.zCommand.SQL.Text := 'update equipos set eOcupado = "Si" where sContrato = :contrato and sIdEquipo = :categoria;';

        connection.zCommand.ParamByName( 'contrato' ).AsString := Contrato;
        connection.zCommand.ParamByName( 'categoria' ).AsString := Categoria;
        connection.zCommand.ExecSQL;
      finally
      end;

    end;

  end;

  procedure TCuadre.DesbloquearCategoria;
  begin
    try
      connection.zCommand.Active := False;

      if Cuadre.TipoCategoria = ftPersonal then
        connection.zcommand.SQL.Text := 'update personal set eOcupado = "No" where sContrato = :contrato and sIdPersonal = :categoria;'
      else
        connection.zcommand.SQL.Text := 'update equipos set eOcupado = "No" where sContrato = :contrato and sIdEquipo = :categoria;';

      connection.zcommand.ParamByName( 'contrato' ).AsString := Contrato;
      connection.zcommand.ParamByName( 'categoria' ).AsString := Categoria;
      connection.zcommand.ExecSQL;
    finally
    end;
  end;

  procedure TCuadre.CargarActividades( var cdActividad, cdFolios : TClientDataSet; Formula : TftFormula );
  var
    zActividades : TZReadOnlyQuery;
  begin

    try
      InicializarZQuery( zActividades );
      zActividades.SQL.Text := SQL_ACTIVIDADES_Y_FOLIOS_REPORTADOS;
      zActividades.ParamByName( 'orden' ).AsString := Cuadre.Orden;
      zActividades.ParamByName( 'fecha' ).AsString := Cuadre.Fecha;
      zActividades.ParamByName( 'x_folio' ).AsInteger := TODAS_LAS_ACTIVIDADES;
      zActividades.ParamByName( 'formula' ).AsInteger := Integer( Formula );
      zActividades.Open;
      zActividades.First;
      cdFolios.First;

      if cdActividad.RecordCount > 0 then
        cdActividad.EmptyDataSet;

      while not cdFolios.Eof do
      begin
        zActividades.Filtered := False;
        zActividades.Filter := 'sNumeroOrden = ' + QuotedStr( cdFolios.FieldByName( 'Folio' ).AsString );
        zActividades.Filtered := True;
        zActividades.First;

        while not zActividades.Eof do
        begin
          cdActividad.Append;
          cdActividad.FieldByName( 'IdRegistro' ).AsInteger := cdActividad.RecordCount + 1;
          cdActividad.FieldByName( 'IdRegistroFolio' ).AsInteger := cdFolios.FieldByName( 'IdRegistro' ).AsInteger;
          cdActividad.FieldByName( 'IdActividad' ).AsInteger := zActividades.FieldByName( 'iIdActividad' ).AsInteger;
          cdActividad.FieldByName( 'Actividad' ).AsString := zActividades.FieldByName( 'sNumeroActividad' ).AsString;
          cdActividad.FieldByName( 'Wbs' ).AsString := zActividades.FieldByName( 'sWbs' ).AsString;
          cdActividad.FieldByName( 'HInicio' ).AsString := zActividades.FieldByName( 'sHoraInicio' ).AsString;
          cdActividad.FieldByName( 'HFin' ).AsString := zActividades.FieldByName( 'sHoraFinal' ).AsString;
          cdActividad.FieldByName( 'Fila' ).AsInteger := FILA_NO_ASIGNADA;
          cdActividad.FieldByName( 'Duracion' ).AsString := zActividades.FieldByName( 'Duracion' ).AsString;
          cdActividad.FieldByName( 'IdDiario' ).AsInteger := zActividades.FieldByName( 'iIdDiario' ).AsInteger; 
          cdActividad.FieldByName( 'Tarea' ).AsInteger := zActividades.FieldByName( 'iIdTarea' ).AsInteger;
          cdActividad.FieldByName( 'Descripcion' ).AsString := zActividades.FieldByName( 'mDescripcion' ).AsString;
          cdActividad.Post;

          zActividades.Next;

        end;

        cdFolios.Next;

      end;

    finally
      zActividades.Filtered := False;
      zActividades.Free;
    end;

  end;

  procedure TCuadre.CargarHorarios(var cdFolio: TClientDataSet; var cdActividad: TClientDataSet; var cdHorarios: TClientDataSet; var cdCategoria : TClientDataSet );
  var
    zBuscaCuadre : TZReadOnlyQuery;
  begin
    try
      zBuscaCuadre := TZReadOnlyQuery.Create( nil );
      zBuscaCuadre.Connection := connection.zConnection;
      zBuscaCuadre.Active := False;
      zBuscaCuadre.SQL.Text := SQL_CUADRE_EXISTENTE[ Integer( Cuadre.TipoCategoria ) ];
      zBuscaCuadre.ParamByName( 'contrato' ).AsString := Contrato;
      zBuscaCuadre.ParamByName( 'orden' ).AsString := Orden;
      zBuscaCuadre.ParamByName( 'fecha' ).AsString := Fecha;
      zBuscaCuadre.ParamByName( 'idrecurso' ).AsString := Cuadre.Categoria;
      zBuscaCuadre.Open;
      zBuscaCuadre.First;

      cdFolio.First;
      cdActividad.First;

      if cdHorarios.RecordCount > 0 then
        cdHorarios.EmptyDataSet;

      while not cdActividad.Eof do
      begin

        zBuscaCuadre.Filtered := False;
        zBuscaCuadre.Filter := 'iIdActividad = ' + cdActividad.FieldByName( 'IdActividad' ).AsString ;
        zBuscaCuadre.Filtered := True;
        zBuscaCuadre.First;

        while not zBuscaCuadre.Eof do
        begin

          cdHorarios.Append;
          cdHorarios.FieldByName( 'IdRegistro' ).AsInteger := cdHorarios.RecordCount + 1;
          cdHorarios.FieldByName( 'IdRegistroActividad' ).AsInteger := cdActividad.FieldByName( 'IdRegistro' ).AsInteger;
          cdHorarios.FieldByName( 'IdRegistroCategoria' ).AsInteger := cdCategoria.FieldByName( 'IdRegistro' ).AsInteger;
          cdHorarios.FieldByName( 'Inicio' ).AsString := zBuscaCuadre.FieldByName( 'sHoraInicio' ).AsString;
          cdHorarios.FieldByName( 'Fin' ).AsString := zBuscaCuadre.FieldByName( 'sHoraFinal' ).AsString;
          cdHorarios.FieldByName( 'Fila' ).AsInteger := FILA_NO_ASIGNADA;
          cdHorarios.FieldByName( 'Duracion' ).AsString := zBuscaCuadre.FieldByName( 'Duracion' ).AsString;
          cdHorarios.FieldByName( 'Cantidad' ).AsInteger := zBuscaCuadre.FieldByName( 'dCantidad' ).AsInteger;
          cdHorarios.FieldByName( 'CantidadHorasHombre' ).AsFloat := zBuscaCuadre.FieldByName( 'dCantHH' ).AsFloat;
          cdHorarios.Post;

          zBuscaCuadre.Next;

        end;

        if zBuscaCuadre.RecordCount = 0 then
        begin
          cdHorarios.Append;
          cdHorarios.FieldByName( 'IdRegistro' ).AsInteger := cdHorarios.RecordCount + 1;
          cdHorarios.FieldByName( 'IdRegistroActividad' ).AsInteger := cdActividad.FieldByName( 'IdRegistro' ).AsInteger;
          cdHorarios.FieldByName( 'IdRegistroCategoria' ).AsInteger := cdCategoria.FieldByName( 'IdRegistro' ).AsInteger;
          cdHorarios.FieldByName( 'Inicio' ).AsString := cdActividad.FieldByName( 'HInicio' ).AsString;
          cdHorarios.FieldByName( 'Fin' ).AsString := cdActividad.FieldByName( 'HFin' ).AsString;
          cdHorarios.FieldByName( 'Fila' ).AsInteger := FILA_NO_ASIGNADA;
          cdHorarios.FieldByName( 'Duracion' ).AsString := cdActividad.FieldByName( 'Duracion' ).AsString;
          cdHorarios.FieldByName( 'Cantidad' ).AsInteger := 0;
          cdHorarios.FieldByName( 'CantidadHorasHombre' ).AsFloat := 0;
          cdHorarios.Post;
        end;

        zBuscaCuadre.Filtered := False;
        cdActividad.Next;

      end;

    finally
      zBuscaCuadre.Filtered := False;
      cdActividad.Filtered := False;
      cdHorarios.Filtered := False;
      zBuscaCuadre.Free;
    end;
  end;

  procedure TCuadre.DefinirEstructuras(var cdFolio: TClientDataSet; var cdActividad: TClientDataSet; var cdCategoria: TClientDataSet; var cdHorario: TClientDataSet);
  begin

    cdFolio.FieldDefs.Add( 'IdRegistro', ftInteger, 0, True );
    cdFolio.FieldDefs.Add( 'Folio', ftString, 255, True );
    cdFolio.FieldDefs.Add( 'Fila', ftInteger, 0, False );
    cdFolio.FieldDefs.Add( 'Inicio', ftInteger, 0, False );
    cdFolio.FieldDefs.Add( 'Fin', ftInteger, 0, False );
    cdFolio.FieldDefs.Add( 'Pernocta', ftString, 255, False );
    cdFolio.FieldDefs.Add( 'Plataforma', ftString, 255, False );
    cdFolio.FieldDefs.Add( 'Suma', ftString, 255, False );
    cdFolio.FieldDefs.Add( 'Ajuste', ftFloat, 0, False );
    cdFolio.CreateDataSet;

    cdActividad.FieldDefs.Add( 'IdRegistro', ftInteger, 0, True );
    cdActividad.FieldDefs.Add( 'IdRegistroFolio', ftInteger, 0, True );
    cdActividad.FieldDefs.Add( 'IdActividad', ftInteger, 0, True );
    cdActividad.FieldDefs.Add( 'Actividad', ftString, 255, True );
    cdActividad.FieldDefs.Add( 'Wbs', ftString, 255, True );
    cdActividad.FieldDefs.Add( 'Inicio', ftInteger, 0, False );
    cdActividad.FieldDefs.Add( 'Fin', ftInteger, 0, False );
    cdActividad.FieldDefs.Add( 'HInicio', ftString, 100, False );
    cdActividad.FieldDefs.Add( 'HFin', ftString, 100, False );
    cdActividad.FieldDefs.Add( 'Fila', ftInteger, 0, False );
    cdActividad.FieldDefs.Add( 'Duracion', ftString, 255, False );
    cdActividad.FieldDefs.Add( 'IdDiario', ftInteger, 0, True );
    cdActividad.FieldDefs.Add( 'Tarea', ftInteger, 0, True );
    cdActividad.FieldDefs.Add( 'Descripcion', ftMemo, 0, True );
    cdActividad.CreateDataSet;

    cdCategoria.FieldDefs.Add( 'IdRegistro', ftInteger, 0, True );
    cdCategoria.FieldDefs.Add( 'IdCategoria', ftString, 255, True );
    cdCategoria.FieldDefs.Add( 'Descripcion', ftMemo, 0, True );
    cdCategoria.FieldDefs.Add( 'Tierra', ftString, 10, True );
    cdCategoria.FieldDefs.Add( 'Solicitado', ftInteger, 0, True );
    cdCategoria.FieldDefs.Add( 'A_Bordo', ftInteger, 0, True );
    cdCategoria.FieldDefs.Add( 'TiempoExtra', ftString, 10, True );
    cdCategoria.FieldDefs.Add( 'Pernocta', ftString, 255, False );
    cdCategoria.FieldDefs.Add( 'Plataforma', ftString, 255, False );
    cdCategoria.FieldDefs.Add( 'Agrupa', ftString, 255, True );
    cdCategoria.CreateDataSet;

    cdHorario.FieldDefs.Add( 'IdRegistro', ftInteger, 0, True );
    cdHorario.FieldDefs.Add( 'IdRegistroActividad', ftInteger, 0, True );
    cdHorario.FieldDefs.Add( 'IdRegistroCategoria', ftInteger, 0, True );
    cdHorario.FieldDefs.Add( 'Inicio', ftString, 255, True );
    cdHorario.FieldDefs.Add( 'Fin', ftString, 255, True );
    cdHorario.FieldDefs.Add( 'Fila', ftInteger, 0, True );
    cdHorario.FieldDefs.Add( 'Duracion', ftString, 255, False );
    cdHorario.FieldDefs.Add( 'Cantidad', ftInteger, 0, True );
    cdHorario.FieldDefs.Add( 'CantidadHorasHombre', ftFloat, 0, True );
    cdHorario.CreateDataSet;

  end;

  procedure TCuadre.CargarCategorias( var cdCategoria : TClientDataSet; TipoRecurso : TftCategoria );
  var
    zCategoria : TZReadOnlyQuery;
  begin
    try
      InicializarZQuery( zCategoria );
      zCategoria.SQL.Text := SQL_CATEGORIAS[ Integer( TipoRecurso ) ];
      zCategoria.ParamByName( 'contrato' ).AsString := Cuadre.Contrato;
      zCategoria.ParamByName( 'orden' ).AsString := Cuadre.Orden;
      zCategoria.ParamByName( 'fecha' ).AsString := Cuadre.Fecha;
      zCategoria.ParamByName( 'todas_categorias' ).AsInteger := UNA_CATEGORIA;
      zCategoria.ParamByName( 'id_recurso' ).AsString := Cuadre.Categoria;
      zCategoria.Open;
      zCategoria.First;

      if cdCategoria.RecordCount > 0 then
        cdCategoria.EmptyDataSet;
      
      while not zCategoria.Eof do
      begin
        cdCategoria.Append;
        cdCategoria.FieldByName( 'IdRegistro' ).AsInteger := cdCategoria.RecordCount + 1;
        cdCategoria.FieldByName( 'IdCategoria' ).AsString := zCategoria.FieldByName( 'sIdRecurso' ).AsString;
        cdCategoria.FieldByName( 'Descripcion' ).AsString := zCategoria.FieldByName( 'sDescripcion' ).AsString;
        cdCategoria.FieldByName( 'Tierra' ).AsString := zCategoria.FieldByName( 'sTierra' ).AsString;
        cdCategoria.FieldByName( 'Solicitado' ).AsInteger := zCategoria.FieldByName( 'iSolicitado' ).AsInteger;
        cdCategoria.FieldByName( 'A_Bordo' ).AsInteger := zCategoria.FieldByName( 'iAbordo' ).AsInteger;
        cdCategoria.FieldByName( 'TiempoExtra' ).AsString := zCategoria.FieldByName( 'sTE' ).AsString;
        cdCategoria.FieldByName( 'Agrupa' ).AsString := zCategoria.FieldByName( 'sAgrupa' ).AsString;
        cdCategoria.Post;

        Cuadre.IdCategoria := cdCategoria.FieldByName( 'IdRegistro' ).AsInteger;        

        if cdCategoria.FieldByName( 'Tierra' ).AsString = 'Si' then
          Cuadre.Formula := ftTierra
        else
        begin
          if cdCategoria.FieldByName( 'TiempoExtra' ).AsString = 'Si' then
            Cuadre.Formula := ftTiempoExtra
          else
            Cuadre.Formula := ftA_Bordo;
        end;

        if TipoRecurso = ftEquipo then
          Cuadre.Formula := ftEq;

        zCategoria.Next;
      end;
    finally
      zCategoria.Free;
    end;
  end;


  function TCuadre.GenerarCuadre( var cdFolios, cdCategorias, cdActividades, cdHorarios : TClientDataSet; var prgF, prgA, prgH : TProgressBar; Formulas : TftFormula ):TFileName;
  var
    Excel,
    Libro,
    Hoja,
    Rango : TExcelInstance;

    Fila : TExcelFila;
    Columna : TExcelColumna;

    Archivo : TFileName;

    IdHorario : TIdRegistroHorario;
    IdActividad : TIdRegistroActividad;
    IdFolio : TIdRegistroFolio;

    BookFolios,
    BookActividades,
    BookHorarios : TBookmark;

  begin

    try

      {$region 'Crear Objetos'}

      {$REGION 'Crear excel'}

      try
        Excel := CreateOleObject('Excel.Application');
        Libro := Excel.Workbooks.Add;

        Libro.Sheets.Add;
        while Libro.Sheets.Count > 1 do
          Libro.Sheets[1].Delete;

        Hoja := Libro.Sheets[1];
        Excel.DisplayAlerts := False;
        Excel.ScreenUpdating := True;
        Excel.Visible :=  False;
        Excel.Workbooks[1].Sheets[1].Name := 'CUADRE';
        Excel.Columns[ 'A:A' ].ColumnWidth := 0;

      except
        on e:Exception do
        begin
          MessageDlg('No se puede continuar verifique tener instalada la suite de Microsoft Office', mtInformation, [mbOK], 0);
          Exit;
        end;
      end;

      {$ENDREGION}

      Fila := TExcelFila.Create;
      Columna := TExcelColumna.Create;

      {$endregion}

      Columna.Column := 2;
      Fila.Row := 2;

      prgF.Max := cdFolios.RecordCount;
      prgF.Position := 0;
      Application.ProcessMessages;

      cdFolios.Filtered := False;
      cdActividades.Filtered := False;
      cdCategorias.Filtered := False;
      cdHorarios.Filtered := False;

      cdFolios.First;
      cdActividades.First;
      cdCategorias.First;
      cdHorarios.First;

      cdFolios.IndexFieldNames := EMPTY_STRING;
      while not cdFolios.Eof do
      begin

        {$region 'Cabecera de la categoria'}

        Rango := Excel.Range[ 'B' + Fila.StrFila + ':C' + Fila.StrFila( 1 ) ];
        Rango.MergeCells := True;
        Rango.Value := UpperCase( Cuadre.Tipo );
        Rango.HorizontalAlignment := xlCenter;
        Rango.VerticalAlignment := xlCenter;

        Rango := Excel.Range[ 'D' + Fila.StrFila + ':D' + Fila.StrFila( 1 ) ];
        Rango.MergeCells := True;
        Rango.Value := Cuadre.Categoria;
        Rango.HorizontalAlignment := xlCenter;
        Rango.VerticalAlignment := xlCenter;

        Rango := Excel.Range[ 'E' + Fila.StrFila + ':I' + Fila.StrFila( 1 ) ];
        Rango.MergeCells := True;
        Rango.Value := cdCategorias.FieldByName( 'Descripcion' ).AsString;
        Rango.HorizontalAlignment := xlJustify;
        Rango.VerticalAlignment := xlCenter;

        Fila.SigFila( 2 );

        Rango := Excel.Range[ 'E' + Fila.StrFila ];
        Rango.Value := 'SOLICITADO';

        Rango := Excel.Range[ 'E' + Fila.StrFila( 1 ) ];
        Rango.Value := cdCategorias.FieldByName( 'Solicitado' ).AsString;

        Rango := Excel.Range[ 'F' + Fila.StrFila ];
        Rango.Value := 'A BORDO';

        Rango := Excel.Range[ 'F' + Fila.StrFila( 1 ) ];
        Rango.Value := cdCategorias.FieldByName( 'A_Bordo' ).AsString;

        Rango := Excel.Range[ 'G' + Fila.StrFila ];
        Rango.Value := 'TOTAL HH';

        Rango := Excel.Range[ 'B' + Fila.StrFila( -2 ) + ':I' + Fila.StrFila( -1 ) ];
        Rango.Interior.ColorIndex := 33;
        Rango := Excel.Range[ 'E' + Fila.StrFila + ':G' + Fila.StrFila( 1 ) ];
        Rango.HorizontalAlignment := xlCenter;
        Rango.VerticalAlignment := xlCenter;
        Rango.Font.Size := 12;
        Rango.Interior.ColorIndex := 42;

        Fila.SigFila( 3 );

        {$endregion}

        {$region 'Cabecera del Folio'}

        Rango := Excel.Range[ 'B'+Fila.StrFila + ':B' + Fila.StrFila( 1 ) ];
        Rango.MergeCells := True;
        Rango.HorizontalAlignment := xlCenter;
        Rango.VerticalAlignment := xlCenter;
        Rango.Value := 'FOLIO';

        Rango := Excel.Range[ 'C'+ Fila.StrFila + ':E' + Fila.StrFila( 1 ) ];
        Rango.MergeCells := True;
        Rango.Font.Bold := True;
        Rango.HorizontalAlignment := xlCenter;
        Rango.VerticalAlignment := xlCenter;
        Rango.Value := cdFolios.FieldByName( 'Folio' ).AsString;

        Rango := Excel.Range[ 'B' + Fila.StrFila + ':C' + Fila.StrFila ];
        Rango.Interior.ColorIndex := 45;

        {$endregion}

        IdFolio := cdFolios.RecNo;
        cdFolios.Edit;
        cdFolios.FieldByName( 'Fila' ).AsInteger := Fila.Row;
        cdFolios.FieldByName( 'Inicio' ).AsInteger := Fila.Row;
        cdFolios.Post;
        cdFolios.RecNo := IdFolio;

        Fila.SigFila( 2 );

        cdActividades.Filtered := False;
        cdActividades.Filter := 'IdRegistroFolio=' + cdFolios.FieldByName( 'IdRegistro' ).AsString ;
        cdActividades.Filtered := True;
        cdActividades.First;

        while ( not cdActividades.Eof ) do
        begin

          prgA.Max := cdActividades.RecordCount;
          prgA.Position := 0;
          Application.ProcessMessages;

          {$region 'Cabecera Actividad'}

          Rango := Excel.Range[ 'B' + Fila.StrFila ];
          Rango.Value := 'ACTIVIDAD';
          Rango := Excel.Range[ 'C' + Fila.StrFila ];
          Rango.Value := cdActividades.FieldByName( 'Actividad' ).AsString;
          Rango := Excel.Range[ 'D' + Fila.StrFila ];
          Rango.Value := cdActividades.FieldByName( 'HInicio' ).AsString;
          Rango := Excel.Range[ 'E' + Fila.StrFila ];
          Rango.Value := cdActividades.FieldByName( 'HFin' ).AsString;

          Rango := Excel.Range[ 'B' + Fila.StrFila + ':E' + Fila.StrFila ];
          Rango.HorizontalAlignment := xlCenter;
          Rango.VerticalAlignment := xlCenter;
          Rango.Interior.ColorIndex := 45;

          {$endregion}

          IdActividad := cdActividades.RecNo;
          cdActividades.Edit;
          cdActividades.FieldByName( 'Fila' ).AsInteger := Fila.Row;
          cdActividades.FieldByName( 'Inicio' ).AsInteger := Fila.Row + 2;
          cdActividades.Post;
          cdActividades.RecNo := IdActividad;

          Fila.SigFila();

          {$region 'Cabecera Horarios'}

          Excel.Range[ 'C' + Fila.StrFila ].Value := 'INICIO';
          Excel.Range[ 'D' + Fila.StrFila ].Value := 'TERMINO';
          Excel.Range[ 'E' + Fila.StrFila ].Value := 'DURACION';
          Excel.Range[ 'F' + Fila.StrFila ].Value := 'CANTIDAD';
          Excel.Range[ 'G' + Fila.StrFila ].Value := 'HH';

          Rango := Excel.Range[ 'C' + Fila.StrFila + ':G' + Fila.StrFila ];
          Rango.HorizontalAlignment := xlCenter;
          Rango.VerticalAlignment := xlCenter;
          Rango.Interior.ColorIndex := 44;
          Rango.Font.Bold := True;

          {$endregion}

          cdHorarios.Filtered := False;
          cdHorarios.Filter := 'IdRegistroActividad=' + cdActividades.FieldByName( 'IdRegistro' ).AsString;
          cdHorarios.Filtered := True;
          cdHorarios.First;

          prgH.Max := cdActividades.RecordCount;
          prgH.Position := 0;
          Application.ProcessMessages;
          Fila.SigFila();
          Excel.Range[ 'F' + Fila.StrFila ].Locked := False;

          while ( not cdHorarios.Eof ) do
          begin

            IdHorario := cdHorarios.RecNo;
            cdHorarios.Edit;
            cdHorarios.FieldByName( 'Fila' ).AsInteger := Fila.Row;
            cdHorarios.Post;
            cdHorarios.RecNo := IdHorario;

            Excel.Range[ 'C' + Fila.StrFila ].Value := cdHorarios.FieldByName( 'Inicio' ).AsString;
            Excel.Range[ 'D' + Fila.StrFila ].Value := cdHorarios.FieldByName( 'Fin' ).AsString;
            Excel.Range[ 'F' + Fila.StrFila ].Value := cdHorarios.FieldByName( 'Cantidad' ).AsInteger;

            if cdHorarios.FieldByName( 'Fin' ).AsString = '24:00' then
              Excel.Range[ 'D' + Fila.StrFila ].NumberFormat := '[h]:mm:ss'
            else
              Excel.Range[ 'D' + Fila.StrFila ].NumberFormat := 'hh:mm';

            Rango := Excel.Range[ 'E' + Fila.StrFila ];
            Rango.Value := cdHorarios.FieldByName( 'Duracion' ).AsString;
            Rango.NumberFormat := '0.0000';

            //Formulas, Tierra, A Bordo, Tiempo Extra, Equipo
            Rango := Excel.Range[ 'G' + Fila.StrFila ];
            Rango.Formula := '=( E' + Fila.StrFila + ' * F' + Fila.StrFila + ' ) * 24';
            if Formulas = ftA_Bordo then
              Rango.Formula := '=( E' + Fila.StrFila + ' * F' + Fila.StrFila + ' ) * 2'
            else
            begin
              if Formulas = ftTierra then
                Rango.Formula := '=( E' + Fila.StrFila + ' * F' + Fila.StrFila + ' ) * 3'
              else
                Rango.Formula := '=( E' + Fila.StrFila + ' * F' + Fila.StrFila + ' )';
            end;

            Excel.Range[ 'F' + Fila.StrFila ].Locked := False;

            cdHorarios.Next;
            Fila.SigFila()
          end;

//          cdHorarios.Filtered := False;
          cdActividades.Edit;
          cdActividades.FieldByName( 'Fin' ).AsInteger := Fila.Row ;
          cdActividades.Post;
          cdActividades.RecNo := IdActividad;
          cdActividades.Next;
          prgA.Position := prgA.Position + 1;
          Fila.SigFila( 4 );
        end;

        prgF.Position := prgF.Position + 1;
        cdActividades.Filtered := False;
        cdFolios.Edit;
        cdFolios.FieldByName( 'Fin' ).AsInteger := Fila.Row - 4;
        cdFolios.Post;
        cdFolios.Next;
        Fila.SigFila( 2 );
      end;

    finally

      cdFolios.Filtered := False;
      cdActividades.Filtered := False;
      cdCategorias.Filtered := False;
      cdHorarios.Filtered := False;

      try
        RegenerarFormulaSuma( Excel, cdFolios, cdActividades );
      finally
        cdHorarios.IndexFieldNames := 'Fila';
      end;


      Libro.ActiveSheet.Protect( True, True, True );
      GetTempPath(SizeOf(global_TempPath), global_TempPath);
      Path := global_TempPath;
      Archivo := global_TempPath+ GenerarNombreAleatorio(  ) +'.xls';
      Libro.SaveAs( Archivo, 56 );
      Excel.Quit;

      Result := Archivo;
    end;


  end;

  function TCuadre.RegenerarCuadre( var cdFolios, cdActividades, cdHorario : TClientDataSet; DirArchivo : TFileName; Row : Integer; HInicio, HFin, Duracion : string; Formulas : TftFormula ):TFileName;
  var
    Excel,
    Libro,
    Hoja,
    Rango : TExcelInstance;

    Fila : TExcelFila;
    Columna : TExcelColumna;
  begin

    try

      {$REGION 'Crear excel'}

      try
        Excel := CreateOleObject('Excel.Application');
        Excel.DisplayAlerts := False;
        Excel.ScreenUpdating := True;
        Excel.Visible := False;
        Excel.Workbooks.Open( DirArchivo );
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

      Fila := TExcelFila.Create;
      Fila.Row := Row;
      Columna := TExcelColumna.Create;

      {$ENDREGION}

      //Insertar la nueva fila.
      Excel.Rows[ Fila.Row ].insert;
      Excel.Range[ 'C' + Fila.StrFila ].Value := HInicio;
      Excel.Range[ 'D' + Fila.StrFila ].Value := HFin;
      if HFin = '24:00' then
        Excel.Range[ 'D' + Fila.StrFila ].NumberFormat := '[h]:mm:ss'
      else
        Excel.Range[ 'D' + Fila.StrFila ].NumberFormat := 'hh:mm';

      Rango := Excel.Range[ 'C' + Fila.StrFila + ':G' + Fila.StrFila ];
      Rango.Interior.ColorIndex := 2;
      Rango := Excel.Range[ 'E' + Fila.StrFila ];
      Rango.Value := Duracion;
      Rango.NumberFormat := '0.0000';
      Excel.Range[ 'F' + Fila.StrFila ].Value := 0; 
      Rango := Excel.Range[ 'G' + Fila.StrFila ];

      //Formulas, Tierra, A Bordo, Tiempo Extra, Equipo
      Rango.Formula := '=( E' + Fila.StrFila + ' * F' + Fila.StrFila + ' ) * 24';
      if Formulas = ftA_Bordo then
        Rango.Formula := '=( E' + Fila.StrFila + ' * F' + Fila.StrFila + ' ) * 2'
      else
      begin
        if Formulas = ftTierra then
          Rango.Formula := '=( E' + Fila.StrFila + ' * F' + Fila.StrFila + ' ) * 3'
        else
          Rango.Formula := '=( E' + Fila.StrFila + ' * F' + Fila.StrFila + ' )';
      end;

      Excel.Range[ 'F' + Fila.StrFila ].Locked := False;


    finally

      RegenerarFormulaSuma( Excel, cdFolios, cdActividades );
      Libro.ActiveSheet.Protect( True, True, True );
      Libro.SaveAs( DirArchivo, 56 );
      Result := DirArchivo;

      Excel.Quit;
    end;

  end;

  function TCuadre.RegenerarFormulaSuma( var ExcelApp : TExcelInstance; cdFolios, cdActividades : TClientDataSet ):TExcelFormula;
  var
    StrFormula : TExcelFormula;
  begin
    cdFolios.First;
    while not cdFolios.Eof do
    begin
      StrFormula := '=SUM( ';
      cdActividades.First;

      if cdActividades.RecordCount > 0 then
      begin

        while not cdActividades.Eof do
        begin
          StrFormula := StrFormula + 'SUM( G' + IntToStr( cdActividades.FieldByName( 'Inicio' ).AsInteger ) + ':G' + IntToStr( cdActividades.FieldByName( 'Fin' ).AsInteger ) + ' ) + ' ;
          cdActividades.Next;
        end;

        StrFormula := Trim( StrFormula );
        StrFormula[ Length( StrFormula ) ] := ')';

        ExcelApp.Range[ 'G' + IntToStr( cdFolios.FieldByName( 'Fila' ).AsInteger - 2 ) ].Formula := StrFormula;

      end;
      cdFolios.Next;

    end;
    cdFolios.First;

  end;

  function TCuadre.GuardarCuadre( var cdFolios, cdActividades, cdCategoria, cdHorarios : TClientDataSet; Ruta : TFileName ):TFileName;
  var
    Fila : TExcelFila;

    Excel,
    Libro,
    Hoja,
    Rango : TExcelInstance;

    Cantidad : Integer;
    HorasHombre : Double;

    zCommand : TZQuery;

    Insertados : Integer;

  begin

    try
      try

        connection.zConnection.StartTransaction;

        {$region 'Crear excel'}

        try
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

        Fila := TExcelFila.Create;
        Insertados := 0;

        {$endregion}

        cdFolios.First;
        cdActividades.First;
        cdCategoria.First;

        {$region 'Vacia la bitacora'}

        zCommand := TZQuery.Create( nil );
        zCommand.Connection := connection.zConnection;
        zCommand.Active := False;
        zCommand.SQL.Text := SQL_LIMPIA_BITACORA[ Integer( Cuadre.TipoCategoria ) ];
        zCommand.ParamByName( 'orden' ).AsString := Cuadre.Orden;
        zCommand.ParamByName( 'fecha' ).AsString := Cuadre.Fecha;
        zCommand.ParamByName( 'categoria' ).AsString := Cuadre.Categoria;
        zCommand.ExecSQL;

        {$endregion}

        while not cdFolios.Eof do
        begin

          Fila.Row := cdFolios.FieldByName( 'Fila' ).AsInteger;
          Cantidad := Excel.Range[ 'G' + Fila.StrFila( -2 ) ].Value;

          if Cantidad > cdCategoria.FieldByName( 'A_Bordo' ).AsInteger then
          begin
            cdFolios.Next;
            Continue;
          end;

          cdActividades.Filtered := False;
          cdActividades.Filter := 'IdRegistroFolio = ' + cdFolios.FieldByName( 'IdRegistro' ).AsString;
          cdActividades.Filtered := True;
          cdActividades.First;

          while not cdActividades.Eof do
          begin

            cdHorarios.Filtered := False;
            cdHorarios.Filter := 'IdRegistroActividad = ' + cdActividades.FieldByName( 'IdRegistro' ).AsString ;
            cdHorarios.Filtered := True;
            cdHorarios.First;

            while not cdHorarios.Eof do
            begin

              Fila.Row := cdHorarios.FieldByName( 'Fila' ).AsInteger;
              Cantidad := Excel.Range[ 'F' + Fila.StrFila ].Value;

              if Cantidad > 0 then
              begin

                with zCommand do
                begin

                  Active := False;
                  SQL.Text := SQL_INSERT[ Integer( Cuadre.TipoCategoria )];
                  ParamByName( 'orden' ).AsString := Cuadre.Orden;
                  ParamByName( 'iddiario' ).AsInteger := cdActividades.FieldByName( 'IdDiario' ).AsInteger;
                  ParamByName( 'ItemOrden' ).AsInteger := 0;
                  ParamByName( 'fecha' ).AsString := Cuadre.Fecha;
                  ParamByName( 'idrecurso' ).AsString := Cuadre.Categoria;
                  ParamByName( 'descripcion' ).AsString := cdCategoria.FieldByName( 'Descripcion' ).AsString;
                  ParamByName( 'pernocta' ).AsString := cdFolios.FieldByName( 'Pernocta' ).AsString;
                  ParamByName( 'hinicio' ).AsString := cdHorarios.FieldByName( 'Inicio' ).AsString;
                  ParamByName( 'hfinal' ).AsString := cdHorarios.FieldByName( 'Fin' ).AsString;
                  ParamByName( 'cantidad' ).AsInteger := Cantidad;
                  ParamByName( 'wbs' ).AsString := cdActividades.FieldByName( 'Wbs' ).AsString;
                  ParamByName( 'actividad' ).AsString := cdActividades.FieldByName( 'Actividad' ).AsString;
                  Params.ParamByName('cantidadhh').DataType := ftFloat;
                  ParamByName( 'cantidadhh' ).Value := Excel.Range[ 'G' + Fila.StrFila ].Value;
                  ParamByName( 'folio' ).AsString := cdFolios.FieldByName( 'Folio' ).AsString;
                  ParamByName( 'idactividad' ).AsInteger := cdActividades.FieldByName( 'IdActividad' ).AsInteger;
                  ParamByName( 'tarea' ).AsInteger := cdActividades.FieldByName( 'Tarea' ).AsInteger;

                  if Cuadre.TipoCategoria = ftPersonal then
                  begin
                    ParamByName( 'plataforma' ).AsString := cdFolios.FieldByName( 'Plataforma' ).AsString;
                    ParamByName( 'Categoria' ).AsString := cdCategoria.FieldByName( 'Agrupa' ).AsString;
                  end;

                  ExecSQL;

                  Excel.Range[ 'G' + Fila.StrFila ].Interior.ColorIndex := 43;
                  Inc( Insertados );

                end;

              end;

              cdHorarios.Next;

            end;

            cdActividades.Next;

          end;

          cdFolios.Next;

        end;

        connection.zConnection.Commit;

      except
        on e:Exception do
        begin
          connection.zConnection.Rollback;
          MessageDlg( 'Ha ocurrido un error al guardar los datos. ' + #13 + #10 + e.Message + #13 + #10 + ' , se revertiran los cambios.', mtInformation, [ mbOK ], 0 );

        end;

      end;


    finally
      zCommand.Free;
      Libro.SaveAs( Ruta, 56 );
      Excel.Quit;
      Application.ProcessMessages;
      cdFolios.Filtered := False;
      cdActividades.Filtered := False;
      cdHorarios.Filtered := False;

      Cambios := False;

      MessageDlg( IntToStr( Insertados ) + ' registros guardados', mtInformation, [ mbOK ], 0 );

    end;

  end;

  function TExcelFila.SigFila( count : TExcelRow = 1 ):TExcelRow;
  begin
    Inc( Row, count );
    Result := Row;
  end;

  function TExcelFila.StrFila:string;
  begin
    Result := IntToStr( Row );
  end;

  function TExcelFila.StrFila( Increment : Integer ):string;
  begin
    Result := IntToStr( Row + Increment );
  end;

  constructor TExcelColumna.Create;
  begin
    inherited Create;
    Column := 1;
    StrColumn := ColumnaNombre( Column );
  end;

  function TExcelColumna.Columna():TExcelColAlias;
  begin
    StrColumn := ColumnaNombre( Column );
    Result := StrColumn;
  end;

  function TExcelColumna._Columna:TExcelColAlias;
  begin
    Result := ColumnaNombre( Column - 1 );
  end;

  function TExcelColumna.Columna_:TExcelColAlias;
  begin
    Result := ColumnaNombre( Column + 1 );
  end;

  function TExcelColumna._Columna( Increment : TExcelColIndex ):TExcelColAlias;
  begin
    Result := ColumnaNombre( Column - Increment );
  end;

  function TExcelColumna.Columna_( Increment : TExcelColIndex ):TExcelColAlias;
  begin
    Result := ColumnaNombre( Column + Increment );
  end;


  procedure InicializarZQuery( var zQuery : TZQuery );overload;
  begin
    zQuery := TZQuery.Create( nil );
    zQuery.Connection := connection.zConnection;
    zQuery.Active := False;
  end;

  procedure InicializarZQuery( var zQuery : TZReadOnlyQuery );
  begin
    zQuery := TZReadOnlyQuery.Create( nil );
    zQuery.Connection := connection.zConnection;
    zQuery.Active := False;
  end;

  procedure MostrarCategorias( var TreeRecursos : TcxTreeView; TipoRecurso : TftCategoria; ConservarDatos : Boolean = True );
  var
    zBusca : TZReadOnlyQuery;
    Categoria : TTreeNode;
  begin
    try
      InicializarZQuery( zBusca );
      zBusca.SQL.Text := SQL_CATEGORIAS[ Integer( TipoRecurso ) ];
      zBusca.ParamByName( 'contrato' ).AsString := Cuadre.Contrato;
      zBusca.ParamByName( 'orden' ).AsString := Cuadre.Orden;
      zBusca.ParamByName( 'fecha' ).AsString := Cuadre.Fecha;
      zBusca.ParamByName( 'todas_categorias' ).AsInteger := TODAS_LAS_CATEGORIAS;
      zBusca.ParamByName( 'id_recurso' ).AsString := '!';
      zBusca.Open;
      zBusca.First;

      if not ConservarDatos then
        TreeRecursos.Items.Clear;

      Categoria := TreeRecursos.Items.AddChild( nil, TIPO_CATEGORIA[ Integer( TipoRecurso ) ] );
      Categoria.ImageIndex := Integer( TipoRecurso );

      while not zBusca.Eof do
      begin

        with TreeRecursos.Items.AddChild( Categoria, zBusca.FieldByName( 'sIdRecurso' ).AsString + ' - ' + zBusca.FieldByName( 'sDescripcion' ).AsString ) do
        begin
          ImageIndex := 2;
          StateIndex := Integer( ftCategoriaLibre );
          if zBusca.FieldByName( 'eOcupado' ).AsString = 'Si' then
            StateIndex := Integer( ftCategoriaBloqueada );
          
        end;

        zBusca.Next;

      end;

      Categoria := nil;      

    finally
      zBusca.Free;

    end;
  end;

  procedure CargarCuadresHechos( var TreeDias : TcxTreeView );
  var
    zDias : TZQuery;
    treeOrden : TTreeNode;
  begin
    try
      InicializarZQuery( zDias );
      zDias.SQL.Text := SQL_CUADRES_EXISTENTES;
      zDias.ParamByName( 'contrato' ).AsString := global_Contrato_Barco;
      zDias.ParamByName( 'orden' ).AsString := global_contrato;
      zDias.Open;
      zDias.First;
    
      TreeDias.Items.Clear;
      treeOrden := TreeDias.Items.AddChildFirst( nil, global_contrato );
      treeOrden.ImageIndex := 3;
    
      while not zDias.Eof do
      begin
        with TreeDias.Items.AddChild( treeOrden, zDias.FieldByName( 'dIdFecha' ).AsString ) do
          ImageIndex := 4;

        zDias.Next;
      end;

      treeOrden := nil;    

    finally
      zDias.Free;
    end;
  end;

  procedure InicializarForm( var Ventana : TForm; Grupo : TcxGroupBox; Largo, Ancho : Integer );
  begin
    Ventana := TForm.Create( nil );
    Ventana.Position := poScreenCenter;
    Ventana.BorderStyle := bsDialog;
    Ventana.Width := Largo;
    Ventana.Height := Ancho;
    Ventana.Caption := EMPTY_STRING;
    Grupo.Parent := Ventana;
    Grupo.Align := alClient;
    Grupo.Visible := True;    
  end;

  procedure ReestablecerBox( var Box : TcxGroupBox; NuevoParent : TForm );
  begin
    Box.Parent := NuevoParent;
    Box.Width := 0;
    Box.Height := 0;
    Box.Left := 0;
    Box.Top := 0;
    Box.Visible := False;
  end;

  function TCuadre.ObtenerDuracion( HFin, HInicio : String; Formulas : TftFormula ):string;
  var
    ZDuracion : TZReadOnlyQuery;
  begin
    try
      ZDuracion := TZReadOnlyQuery.Create( nil );
      ZDuracion.Connection := connection.zConnection;

      with ZDuracion do
      begin
        Active := False;
        SQL.Text := 'select time_format( timediff( :Final, :Inicio ), if( :formula = 2, "%H.%i", "%H:%i" ) ) as Duracion;';
        ParamByName( 'Final' ).AsString := HFin;
        ParamByName( 'Inicio' ).AsString := HInicio;
        ParamByName( 'formula' ).AsInteger := Integer( Formulas );
        Open;
      end;

      Result := ZDuracion.FieldByName( 'Duracion' ).AsString;

    finally
      ZDuracion.Free;
    end;
  end;

end.
