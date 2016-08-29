unit UnitTarifa_Calidad;

interface

uses frxClass,Forms, frxDBSet,ZConnection, DB, ZAbstractRODataset, Dialogs,
  ZAbstractDataset, ZDataset,math,strUtils,DateUtils,SysUtils, RxMemDS,
  classes,JvMemoryDataset,DBClient,AdvProgr,cxProgressBar,ComObj,OleCtrls,
  UFunctionsGHH;

Type
  FtTipo=(FtTierra,FtAbordo);
  FtSeccion=(FtsRIR,FtsRIS,FtsAll, FtsNone);

  Procedure EncabezadoPDF_Horizontal(Reportediario :TzREadOnlyQuery; Var Reporte: TfrxReport;Tipo:FtTipo);
  Procedure FirmasPDF_Generales(Reportediario :TzREadOnlyQuery;      Var Reporte: TfrxReport;Tipo:FtTipo);
  Procedure ReporteCalidad_RIR(ReporteDiario:TzREadOnlyQuery;        Var Reporte: TfrxReport;Tipo:FtTipo;TImpresion:FtSeccion);

  const
    TotalCol=7;
    TotalColPer=5;
    Mantisa=2;

implementation

uses frm_connection, global,Utilerias;

procedure ReportePDF_ClearDataset(Var Reporte: TfrxReport);
var
  i:Integer;
begin
  for I := 0 to Reporte.DataSets.Count - 1 do
    Reporte.DataSets.Items[i].DataSet.Destroy;
  Reporte.DataSets.Clear;
end;

Procedure ReporteCalidad_RIR(ReporteDiario:TzREadOnlyQuery;Var Reporte: TfrxReport;Tipo:FtTipo;TImpresion:FtSeccion);

Var
  i, x,y,CuentaCol,iGrupo,iCiclo: Integer;
  DCiclo,dTotal:Double;
  QryConsulta,
  QrMoe,
  QrColumnas,
  QrRecursos,
  QryCondiciones,
  QryEmbarcacion,
  QryOtroPersonal: TZQuery;
  QryAgrupador,
  qrReportado,
  qrPernoctas : TZQuery;

  dContrato_Inicio,
  dContrato_Final: TDateTime;
  mDatos:TJvMemoryData;
  mImprime:TJvMemoryData;
  Td_ImpDistribucion_detalle: TfrxDBDataset;
  Td_Distribucion_detalle: TfrxDBDataset;
  QrAdicional:TzReadOnlyquery;
  TmpAnexo:String;
  TmpDescAnexo:string;
  ValTmp:variant;
  CantTmp:Double;
  iPosTmp:Integer;
begin

  mDatos:=TJvMemoryData.Create(nil);
  mImprime:=TJvMemoryData.Create(nil);
  QrRecursos  := TZQuery.Create(nil);
  QrColumnas  := TZQuery.Create(nil);
  qrReportado := TZQuery.Create(nil);
  qrPernoctas := TZQuery.Create(nil);
  QrAdicional := TzReadOnlyquery.Create(nil);
  QrMoe       := TZQuery.Create(nil);
  Td_Distribucion_detalle:=TfrxDBDataset.Create(nil);
  Td_ImpDistribucion_detalle:=TfrxDBDataset.Create(nil);
  try
    QrMoe.Connection := Connection.zConnection;
    QrColumnas.Connection  := Connection.zConnection;
    QrRecursos.Connection  := Connection.zConnection;
    QrReportado.Connection := Connection.zConnection;
    QrPernoctas.Connection := Connection.zConnection;
    QrAdicional.Connection := Connection.zConnection;
    Td_Distribucion_detalle.UserName    :='Td_Distribucion_detalle';
    Td_ImpDistribucion_detalle.UserName :='Td_ImpDistribucion_detalle';
    with mDatos do
    begin
      Active:=false;
      FieldDefs.Add('iGrupo', ftInteger, 0, True);
      FieldDefs.Add('sidAnexo', ftString, 10, false);
      FieldDefs.Add('sidrecurso', ftString, 100, True);
      FieldDefs.Add('sdescripcion', ftString, 250, True);
      FieldDefs.Add('sAnexo', ftString, 250, false);
      FieldDefs.Add('sTitulo', ftString, 250, false);
      FieldDefs.Add('smedida', ftString, 100, True);
      FieldDefs.Add('sTipo', ftString, 10, false);
      FieldDefs.Add('dcantSol', FtFloat, 0, True);
      FieldDefs.Add('dcantTotal', FtFloat, 0, True);
      for CuentaCol:=1 to TotalCol do
      begin
        FieldDefs.Add('dcantidad' + Inttostr(CuentaCol), FtFloat, 0, false);
        FieldDefs.Add('sNumeroOrden'+ Inttostr(CuentaCol), ftString, 100, false);
        FieldDefs.Add('sPernocta' + Inttostr(CuentaCol), ftString, 100, false);
        FieldDefs.Add('sPlataforma'+ Inttostr(CuentaCol), ftString, 100, false);
      end;
      Active:=true;
    end;
    with mImprime do
    begin
      Active:=false;
      FieldDefs.Add('iCampo', ftInteger, 0, True);
      Active:=true;
    end;

    mImprime.EmptyTable;
    if (TImpresion=FtsRIR) then
    begin

        {$REGION 'BARCO'}

        QrColumnas.active:=false;
        QrColumnas.SQL.Add( 'SELECT ot.sContrato, ot.sNumeroOrden, p.sDescripcion AS pernocta, p.sDescripcion AS pernocta, pf.sDescripcion AS plataforma, '+
                            'bp.sIdpernocta AS PernoctaP, bp.sIdPlataforma AS idPlataforma ' +
                            'FROM ordenesdetrabajo ot ' +
                            'INNER JOIN contratos AS c ON (ot.sContrato=c.sContrato) ' +
                            'INNER JOIN bitacoradepersonal AS bp ON (bp.scontrato=:OT AND bp.sNumeroOrden = ot.sNumeroOrden ) ' +
                            'INNER JOIN pernoctan AS p ON (p.sidPernocta=bp.sIdpernocta) ' +
                            'INNER JOIN plataformas AS pf ON (pf.sidPlataforma=bp.sIdPlataforma) ' +
                            'WHERE (c.sContrato=:OT OR c.sCodigo=:OT)	AND bp.dIdFecha= :Fecha ' +
                            'GROUP BY	ot.sContrato, ot.sNumeroorden, bp.sidPernocta, bp.sIdPlataforma');
        QrColumnas.ParamByName('OT').AsString       := ReporteDiario.FieldByName('sOrden').AsString;
        QrColumnas.ParamByName('Fecha').AsDate      := Reportediario.FieldByName('dIdFecha').AsDateTime;
        QrColumnas.Open;

        qrMoe.Active:=false;
        QrMoe.SQL.Clear;
        qrMoe.SQL.Text:='select  t.sIdTipoMovimiento as sIdRecurso, t.*, a.sanexo,ifnull(sum(sFactor),0) as TotalFactor,a.sdescripcion as anexo, a.stitulo as tituloAnexo '+
                              'from tiposdemovimiento t '+
                              'inner join movimientosdeembarcacion m '+
                              'on (m.sContrato = t.sContrato and m.dIdFecha =:Fecha and m.sClasificacion = t.sIdTipoMovimiento and (m.sIdFase = "OPER" or m.sIdFase ="ESP")) '+
                              'left join anexos a on(a.sTipo= "BARCO") '+
                              'where t.sContrato =:contrato and t.sClasificacion = "Movimiento de Barco" group by m.sClasificacion order by t.iorden';
        qrMoe.ParamByName('Contrato').AsString := global_contrato_barco;
        qrMoe.ParamByName('Fecha').AsDate      := Reportediario.FieldByName('dIdFecha').AsDateTime;
        qrMoe.Open;

        qrRecursos.Active:=false;
        qrRecursos.SQL.Text:= 'select ifnull(sum(mf.sFactor),0) as dCantidad from movimientosxfolios mf' + #10 +
                              'inner join movimientosdeembarcacion me' + #10 +
                              'on(me.sContrato=mf.sContrato and me.dIdFecha=mf.dIdFecha and' + #10 +
                              'me.iIdDiario=mf.iIdDiario and me.sOrden=mf.sNumeroOrden)' + #10 +
                              'where mf.sContrato=:contratoBarco and mf.didfecha=:fecha and' + #10 +
                              'mf.sNumeroOrden=:contrato and mf.sFolio=:folio and me.sClasificacion=:Tipo';

        //movimientosxfolios where sContrato=:Contrato and dIdFecha=:fecha and iIdDiario=:Diario

        DCiclo:=QrColumnas.RecordCount/TotalCol;
        iCiclo:=Trunc(DCiclo);
        if (dCiclo -iCiclo)>0 then
            Inc(iCiclo,1);
        iGrupo:=1;
        while iGrupo<=iCiclo do
        begin
            with qrMoe do
            begin
                first;
                while not eof do
                begin
                    mDatos.Append;
                    mDatos.FieldByName('iGrupo').AsInteger:=Igrupo;
                    mDatos.FieldByName('sidAnexo').AsString     :=FieldByName('sanexo').asstring;
                    mDatos.FieldByName('sAnexo').AsString       :=FieldByName('anexo').asstring;
                    mDatos.FieldByName('sTitulo').AsString      :=FieldByName('tituloAnexo').asstring;
                    mDatos.FieldByName('sidrecurso').AsString   :=FieldByName('sidrecurso').asstring;
                    mDatos.FieldByName('sdescripcion').AsString :=fieldbyname('sdescripcion').asstring;
                    mDatos.FieldByName('smedida').AsString      :=fieldbyname('smedida').asstring;
                    mDatos.FieldByName('dcantSol').AsFloat      :=fieldbyname('dcantidad').AsFloat;
                    mDatos.FieldByName('sTipo').AsString        :='';
                    dTotal:=0;
                    if iGrupo=1 then
                       QrColumnas.First
                    else
                       QrColumnas.RecNo:=((iGrupo-1) * TotalCol)+ 1;

                    CuentaCol:=1;
                    while not (QrColumnas.Eof) and (QrColumnas.RecNo<=((iGrupo) * TotalCol)) do
                    begin
                        mDatos.FieldByName('sNumeroOrden'+ Inttostr(CuentaCol)).AsString := qrColumnas.FieldbyName('sNumeroOrden').AsString;
                        mDatos.FieldByName('sPernocta' + Inttostr(CuentaCol)).AsString   := QrColumnas.Fieldbyname('Pernocta').asstring;
                        mDatos.FieldByName('sPlataforma'+ Inttostr(CuentaCol)).AsString  := QrColumnas.Fieldbyname('Plataforma').Asstring;
                       // mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat:=QrRecursos.FieldByName('dCantidad').AsFloat;

                        QrRecursos.Active := False;
                        //QrRecursos.ParamByName('Equipo').AsString        := FieldByName('sIdRecurso').AsString;
                        QrRecursos.ParamByName('contratoBarco').AsString := global_contrato_barco;
                        QrRecursos.ParamByName('contrato').AsString      := ReporteDiario.FieldByName('sOrden').AsString;
                        QrRecursos.ParamByName('folio').AsString         := QrColumnas.FieldByName('sNumeroOrden').AsString;
                        QrRecursos.ParamByName('fecha').AsDate           := reportediario.FieldByName('dIdFecha').AsDateTime;
                        QrRecursos.ParamByName('tipo').AsString          := FieldByName('sIdRecurso').AsString;
                        QrRecursos.Open;

                        if QrRecursos.RecordCount>0 then
                        begin
                          mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat:=xRound(QrRecursos.FieldByName('dCantidad').Value,4);
                          dTotal:=dTotal+ QrRecursos.FieldByName('dCantidad').AsFloat;
                        end;
                        QrColumnas.next;
                        Inc(CuentaCol);
                    end;

                    //if dTotal>0 then
                    //begin
                        mDatos.FieldByName('dcantTotal').AsFloat:=fieldbyname('TotalFactor').AsFloat;
                        mDatos.Post;
                    //end
                    //else
                    //   mDatos.cancel;
                    next;
                end;
            end;
            Inc(iGrupo,1);
        end;


        {$ENDREGION}

        {$REGION 'PERSONAL DE TERRA Y A BORDO'}
        QrRecursos.Active := False;
        QrRecursos.SQL.Clear;
        QrRecursos.SQL.Add('SELECT bp.sIdPersonal, bp.sDescripcion, ifnull(SUM(bp.dAjuste),0) AS Ajuste, ' +
                           'if(:Anexo=1,( ' +
                           'IF(SUM(bp.dCanthh) > 0, SUM(bp.dCanthh), SUM(bp.dCantidad))'+
                           '),SUM(bp.dCantidad)) AS Total ' +
                           'FROM bitacoradepersonal bp ' +
                           'WHERE bp.scontrato = :Orden AND bp.sNumeroOrden = :Folio AND bp.didfecha = :Fecha ' +
                           'AND bp.sidPernocta = :Pernocta AND bp.sidplataforma = :Plataforma ' +
                           'and bp.sIdPersonal=:Personal GROUP BY bp.sIdPersonal ');


        {$REGION 'CONSULTAS - PARTIDAS'}
        QrColumnas.active:=false;
        QrColumnas.SQL.Clear;
        QrColumnas.SQL.Add( 'SELECT ot.sContrato, ot.sNumeroOrden, p.sDescripcion AS pernocta, p.sDescripcion AS pernocta, pf.sDescripcion AS plataforma, '+
                            'bp.sIdpernocta AS PernoctaP, bp.sIdPlataforma AS idPlataforma ' +
                            'FROM ordenesdetrabajo ot ' +
                            'INNER JOIN contratos AS c ON (ot.sContrato=c.sContrato) ' +
                            'INNER JOIN bitacoradepersonal AS bp ON (bp.scontrato=:OT AND bp.sNumeroOrden = ot.sNumeroOrden ) ' +
                            'INNER JOIN pernoctan AS p ON (p.sidPernocta=bp.sIdpernocta) ' +
                            'INNER JOIN plataformas AS pf ON (pf.sidPlataforma=bp.sIdPlataforma) ' +
                            'WHERE (c.sContrato=:OT OR c.sCodigo=:OT)	AND bp.dIdFecha= :Fecha ' +
                            'GROUP BY	ot.sContrato, ot.sNumeroorden, bp.sidPernocta, bp.sIdPlataforma');
        QrColumnas.ParamByName('OT').AsString       := ReporteDiario.FieldByName('sOrden').AsString;
        QrColumnas.ParamByName('Fecha').AsDate      := reportediario.FieldByName('dIdFecha').AsDateTime;
        QrColumnas.Open;
        {$ENDREGION}


        {$REGION 'CONSULTAS - TODO EL PERSONAL SOLICITADO QUE SE REGISTRA EN EL MOE'}
        QrMoe.Active := False;
        QrMoe.SQL.Clear;   //ifnull(a.sdescripcion,"SIN ANEXO MAR/TIERRA") as anexo
        QrMoe.SQL.Add('SELECT p.sDescripcion,	mr.*, p.lSumaSolicitado, p.sMedida, ' +
                      'a.sanexo,ifnull(a.sdescripcion,"SIN ANEXO MAR/TIERRA") as anexo, a.stitulo as tituloAnexo,a.stierra '+
                      'FROM moe AS m ' +
                      'left JOIN moerecursos AS mr ON (mr.iidMoe=m.iidMoe) ' +
                      'INNER JOIN personal AS p ON (p.scontrato=:Contrato AND p.sidpersonal=mr.sidRecurso) ' +
                      'left join anexos a on(a.sAnexo=p.sAnexo) ' +
                      'WHERE m.didfecha = (SELECT max(didfecha) FROM moe WHERE didfecha <=:Fecha AND sContrato = :OT) ' +
                      'AND m.sContrato = :OT AND mr.eTipoRecurso = "Personal" ORDER BY a.iOrden,p.iItemOrden');
        QrMoe.ParamByName('Contrato').AsString := Global_Contrato_Barco;
        QrMoe.ParamByName('OT').AsString       := ReporteDiario.FieldByName('sOrden').AsString;
        QrMoe.ParamByName('Fecha').AsDateTime  := reportediario.FieldByName('dIdFecha').AsDateTime;
        QrMoe.Open;
        {$ENDREGION}

        {$REGION 'INSERTA EL MOE'}


        DCiclo:=QrColumnas.RecordCount/TotalCol;
        iCiclo:=Trunc(DCiclo);
        if (dCiclo -iCiclo)>0 then
          Inc(iCiclo,1);

        iGrupo:=1;
        TmpAnexo:='';
        TmpDescAnexo:='';
        while iGrupo<=iCiclo do
        begin
            with qrMoe do
            begin

                first;

                while not eof do
                begin
                    if (TmpAnexo='') and (FieldByName('sanexo').asstring<>'') then
                    begin
                      TmpAnexo:=FieldByName('sanexo').asstring;
                      TmpDescAnexo:=FieldByName('anexo').asstring;
                    end;

                    mDatos.Append;
                    mDatos.FieldByName('iGrupo').AsInteger      :=Igrupo;
                    mDatos.FieldByName('sidAnexo').AsString     :=TmpAnexo;//FieldByName('sanexo').asstring;
                    mDatos.FieldByName('sAnexo').AsString       :=TmpDescAnexo;//FieldByName('anexo').asstring;
                    mDatos.FieldByName('sTitulo').AsString      :=FieldByName('tituloAnexo').asstring;
                    mDatos.FieldByName('sidrecurso').AsString   :=FieldByName('sidrecurso').asstring;
                    mDatos.FieldByName('sdescripcion').AsString :=fieldbyname('sdescripcion').asstring;
                    mDatos.FieldByName('smedida').AsString      :=fieldbyname('smedida').asstring;
                    mDatos.FieldByName('dcantSol').AsFloat      :=fieldbyname('dcantidad').AsFloat;
                    mDatos.FieldByName('sTipo').AsString        :='';
                    dTotal:=0;
                    if iGrupo=1 then
                      QrColumnas.First
                    else
                      QrColumnas.RecNo:=((iGrupo-1) * TotalCol)+ 1;

                    CuentaCol:=1;
                    ValTmp:=0;
                    CantTmp:=0;
                    iPosTmp:=0;
                    while not (QrColumnas.Eof) and (QrColumnas.RecNo<=((iGrupo) * TotalCol)) do
                    begin
                        mDatos.FieldByName('sNumeroOrden'+ Inttostr(CuentaCol)).AsString :=qrColumnas.FieldbyName('sNumeroOrden').AsString;
                        mDatos.FieldByName('sPernocta'   + Inttostr(CuentaCol)).AsString :=QrColumnas.Fieldbyname('Pernocta').asstring;
                        mDatos.FieldByName('sPlataforma' + Inttostr(CuentaCol)).AsString :=QrColumnas.Fieldbyname('Plataforma').Asstring;

                        QrRecursos.Active := False;
                        {if fieldbyname('sTierra').asstring='Si' then
                          QrRecursos.ParamByName('Anexo').AsINteger :=0
                        else }
                        QrRecursos.ParamByName('Anexo').AsINteger :=1;
                        QrRecursos.ParamByName('Orden').AsString      := QrColumnas.FieldByName('sContrato').AsString;
                        QrRecursos.ParamByName('Folio').AsString      := QrColumnas.FieldByName('sNumeroOrden').AsString;
                        QrRecursos.ParamByName('Fecha').AsDate        := reportediario.FieldByName('dIdFecha').AsDateTime;
                        QrRecursos.ParamByName('Pernocta').AsString   := QrColumnas.FieldByName('PernoctaP').AsString;
                        QrRecursos.ParamByName('Plataforma').AsString := QrColumnas.FieldByName('idPlataforma').AsString;
                        QrRecursos.ParamByName('Personal').AsString   := FieldByName('sIdRecurso').AsString;
                        QrRecursos.Open;

                        if QrRecursos.RecordCount>0 then
                        begin
                          mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat:=(xRound(QrRecursos.FieldByName('Total').Value,2) + QrRecursos.FieldByName('Ajuste').Value);
                          dTotal:=dTotal+ (xRound(QrRecursos.FieldByName('Total').Value,2) + QrRecursos.FieldByName('Ajuste').Value);
                          ValTmp:= ValTmp + (QrRecursos.FieldByName('Total').Value + QrRecursos.FieldByName('Ajuste').Value);
                          if CantTmp<mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat then
                          begin
                            CantTmp:=mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat;
                            iPosTmp:=CuentaCol;
                          end;
                        end;
                        QrColumnas.next;
                        Inc(CuentaCol);
                    end;

                    if dTotal>=0 then
                    begin
                      if dTotal<>xRound(ValTmp,2) then
                      begin
                        mDatos.FieldByName('dcantidad' + Inttostr(iPosTmp)).AsFloat:=mDatos.FieldByName('dcantidad' + Inttostr(iPosTmp)).AsFloat
                          + (xRound(ValTmp,2)-dTotal);
                        mDatos.FieldByName('dcantTotal').AsFloat:=xRound(ValTmp,2);
                      end
                      else
                        mDatos.FieldByName('dcantTotal').AsFloat:=dTotal;
                      mDatos.Post;
                    end
                    else
                      mDatos.Cancel;
                    next;
                end;
            end;
            Inc(iGrupo,1);
        end;

        {$ENDREGION}
        {$ENDREGION}

        {$REGION 'EQUIPOS..'}
        {Equipos...}
        QrRecursos.Active := False;
        QrRecursos.SQL.Clear;
        QrRecursos.SQL.Add( 'select be.sIdEquipo, sum(be.dCantHH) as dCantidad '+
                            'from bitacoradeequipos be '+
                            'inner join equipos e '+
                            'on ( e.sIdEquipo = be.sIdEquipo ) '+
                            'inner join bitacoradeactividades ba '+
                            'on ( ba.sContrato = :contrato and ba.dIdFecha = :fecha '+
                                 'and ba.sNumeroOrden = :folio '+
                                 'and ba.iIdDiario = be.iIdDiario '+
                            'and ba.iIdTarea = be.iIdTarea and ba.iIdActividad = be.iIdActividad) '+
                            'where e.sContrato = :contratoBarco '+
                            'and be.sContrato = :contrato '+
                            'and be.sNumeroOrden = :folio '+
                            'and be.sIdPernocta = :pernocta '+
                            'and be.dIdFecha = :fecha '+
                            'and be.sIdEquipo=:equipo ' +
                            'group by e.sIdEquipo '+
                            'order by e.iItemOrden');

        {$REGION 'CONSULTAS - PARTIDAS'}
        QrColumnas.Active:=false;
        QrColumnas.SQL.Clear;
        QrColumnas.SQL.Add( 'select ot.sContrato,ot.sIdFolio, '+
                                     'ot.sNumeroOrden, '+
                                     'be.sIdPernocta AS PernoctaP, '+
                                     'ot.sIdPlataforma AS idPlataforma, '+
                                     'p.sDescripcion as Pernocta, '+
                                     'pt.sDescripcion as Plataforma '+
                              'from ordenesdetrabajo ot '+
                              'inner join contratos c '+
                              'on ( c.sContrato = ot.sContrato ) '+
                              'inner join bitacoradeequipos be '+
                              'on ( be.sContrato = ot.sContrato and be.sNumeroOrden = ot.sNumeroOrden ) '+
                              'inner join pernoctan p '+
                              'on ( ot.sIdPernocta = p.sIdPernocta) '+
                              'inner join plataformas pt '+
                              'on ( ot.sIdPlataforma = pt.sIdPlataforma ) '+
                              'where c.sContrato = :contrato '+
                              'and be.dIdFecha = :fecha '+
                              'group by ot.sIdFolio, p.sIdPernocta' );
        QrColumnas.ParamByName('Contrato').AsString:= ReporteDiario.FieldByName('sOrden').AsString;
        QrColumnas.ParamByName('Fecha').AsDate:=reportediario.FieldByName('dIdFecha').AsDateTime;
        QrColumnas.Open;
        {$ENDREGION}


        {$REGION 'CONSULTAS - TODOS LOS EQUIPOS REGISTRADOS EN MOE'}

        QrMoe.Active := False;
        QrMoe.SQL.Clear;
        QrMoe.SQL.Add('select mr.sIdRecurso, e.sDescripcion, e.sMedida, mr.dCantidad, '+
                      'a.sanexo,a.sdescripcion as anexo, a.stitulo as tituloAnexo '+
                        'from moe m '+
                        'inner join moerecursos mr '+
                        'on ( mr.iIdMoe = m.iIdMoe ) '+
                        'inner join equipos e '+
                        'on ( e.sContrato = :contratobarco and e.sIdEquipo = mr.sIdRecurso ) '+
                        'left join anexos a on(a.sTipo= "EQUIPO") ' +
                        'where m.dIdFecha = (select max(didfecha) from moe where didfecha <=:Fecha and sContrato = :contrato) '+
                        'and m.sContrato = :contrato '+
                        'and mr.eTipoRecurso = "Equipo" '+
                        'order by e.iItemOrden');
        QrMoe.ParamByName('contratobarco').AsString := Global_Contrato_Barco;
        QrMoe.ParamByName('contrato').AsString := ReporteDiario.FieldByName('sOrden').AsString;
        QrMoe.ParamByName('Fecha').AsDateTime := reportediario.FieldByName('dIdFecha').AsDateTime;
        QrMoe.Open;

        {$ENDREGION}

        {$REGION 'INSERCION DE DATOS - INFORMACIÓN DEL EQUIPO'}

        DCiclo:=QrColumnas.RecordCount/TotalCol;
        iCiclo:=Trunc(DCiclo);
        if (dCiclo -iCiclo)>0 then
            Inc(iCiclo,1);
        iGrupo:=1;
        while iGrupo<=iCiclo do
        begin
            with qrMoe do
            begin
                first;
                while not eof do
                begin
                    mDatos.Append;
                    mDatos.FieldByName('iGrupo').AsInteger       :=Igrupo;
                    mDatos.FieldByName('sidAnexo').AsString      :=FieldByName('sanexo').asstring;
                    mDatos.FieldByName('sAnexo').AsString        :=FieldByName('anexo').asstring;
                    mDatos.FieldByName('sTitulo').AsString       :=FieldByName('tituloAnexo').asstring;
                    mDatos.FieldByName('sidrecurso').AsString    :=FieldByName('sidrecurso').asstring;
                    mDatos.FieldByName('sdescripcion').AsString  :=fieldbyname('sdescripcion').asstring;
                    mDatos.FieldByName('smedida').AsString       :=fieldbyname('smedida').asstring;
                    mDatos.FieldByName('dcantSol').AsFloat       :=fieldbyname('dcantidad').AsFloat;
                     mDatos.FieldByName('sTipo').AsString        :='';
                    dTotal:=0;
                    if iGrupo=1 then
                       QrColumnas.First
                    else
                       QrColumnas.RecNo:=((iGrupo-1) * TotalCol)+ 1;

                    CuentaCol:=1;
                    ValTmp:=0;
                    CantTmp:=0;
                    iPosTmp:=0;
                    while not (QrColumnas.Eof) and (QrColumnas.RecNo<=((iGrupo) * TotalCol)) do
                    begin
                        mDatos.FieldByName('sNumeroOrden'+ Inttostr(CuentaCol)).AsString := qrColumnas.FieldbyName('sNumeroOrden').AsString;
                        mDatos.FieldByName('sPernocta' + Inttostr(CuentaCol)).AsString   := QrColumnas.Fieldbyname('Pernocta').asstring;
                        mDatos.FieldByName('sPlataforma'+ Inttostr(CuentaCol)).AsString  := QrColumnas.Fieldbyname('Plataforma').Asstring;

                        QrRecursos.Active := False;
                        QrRecursos.ParamByName('Equipo').AsString        := FieldByName('sIdRecurso').AsString;
                        QrRecursos.ParamByName('contratoBarco').AsString := global_contrato_barco;
                        QrRecursos.ParamByName('contrato').AsString      := ReporteDiario.FieldByName('sOrden').AsString;
                        QrRecursos.ParamByName('folio').AsString         := QrColumnas.FieldByName('sNumeroOrden').AsString;
                        QrRecursos.ParamByName('fecha').AsDate           := reportediario.FieldByName('dIdFecha').AsDateTime;
                        QrRecursos.ParamByName('pernocta').AsString      := QrColumnas.FieldByName('PernoctaP').AsString;
                        QrRecursos.Open;
                        //ABBY
                        if QrRecursos.RecordCount>0 then
                        begin
                          mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat:=xRound(QrRecursos.FieldByName('dCantidad').Value,2);
                          dTotal:=dTotal+ xRound(QrRecursos.FieldByName('dCantidad').Value,2);
                          ValTmp:=ValTmp + QrRecursos.FieldByName('dCantidad').Value;
                          if CantTmp<mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat then
                          begin
                            CantTmp:=mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat;
                            iPosTmp:=CuentaCol;
                          end;
                        end;
                        QrColumnas.next;
                        Inc(CuentaCol);
                    end;

                    if dTotal>0 then
                    begin
                      if dTotal<>xRound(ValTmp,2) then
                      begin
                        mDatos.FieldByName('dcantidad' + Inttostr(iPosTmp)).AsFloat:=mDatos.FieldByName('dcantidad' + Inttostr(iPosTmp)).AsFloat
                        + (xRound(ValTmp,2)-dTotal);
                        mDatos.FieldByName('dcantTotal').AsFloat:=xRound(ValTmp,2);
                      end
                      else
                        mDatos.FieldByName('dcantTotal').AsFloat:=dTotal;
                      mDatos.Post;
                    end
                    else
                       mDatos.cancel;
                    next;
                end;
            end;
            Inc(iGrupo,1);
        end;


     {Esta secccion es para mostrar la hoja vacia sino existen datos.}
//     if connection.configuracion.FieldValues['eHojasBlanco'] = 'Si' then
//        if MDatos.RecordCount>=0 then
//     else
//        if MDatos.RecordCount>0 then
//     begin
//
//     end;

        mImprime.Append;
        mImprime.FieldByName('iCampo').AsInteger:=1;
        mImprime.Post;

        {$ENDREGION}
        {$ENDREGION}

        {$REGION 'PERNOCTAS..'}
        with QrColumnas do
        begin
          active := false;
          sql.text := 'select ot.sIdFolio, '+
                             'ot.sNumeroOrden, '+
                             'ot.sIdPernocta, '+
                             'ot.sIdPlataforma, '+
                             'p.sDescripcion as sPernocta, '+
                             'pt.sDescripcion as sPlataforma '+
                      'from ordenesdetrabajo ot '+
                      'inner join contratos c '+
                      'on ( c.sContrato = ot.sContrato ) '+
                      'inner join bitacoradeequipos be '+
                      'on ( be.sContrato = ot.sContrato and be.sNumeroOrden = ot.sNumeroOrden ) '+
                      'inner join pernoctan p '+
                      'on ( ot.sIdPernocta = p.sIdPernocta) '+
                      'inner join plataformas pt '+
                      'on ( ot.sIdPlataforma = pt.sIdPlataforma ) '+
                      'where c.sContrato = :contrato '+
                      'and be.dIdFecha = :fecha '+
                      'group by ot.sIdFolio, p.sIdPernocta';
          parambyname('contrato').asstring := ReporteDiario.fieldbyname('sOrden').asstring;
          parambyname('fecha').asdate      := reportediario.fieldbyname('didfecha').asdatetime;
          open;
        end;

        with qrPernoctas do
        begin
          active := false;
          sql.text := 'select c.sIdPernocta, '+
                             'c.sDescripcion, '+
                             'c.sMedida, '+

                       '( sum( bp.dCantHH ) - ifnull(( select ifnull( sum(bpr.dCantidad), 0) '+
                                               'from bitacoradepernocta bpr '+
                                               'where bpr.sContrato = :contrato '+
                                               'and bpr.dIdFecha = :fecha '+
                                               'and (bpr.sNumeroOrden <> "@" and bpr.sNumeroOrden=ba.snumeroorden) group by bpr.dIdFecha),0) ) as dCantidad '+

                      'from cuentas c '+
                      'left join bitacoradepersonal bp '+
                        'on ( '+ //bp.lAplicaPernocta = "Si"
                          'bp.sContrato = :contrato '+
                          'and bp.dIdFecha = :fecha '+
                          'and bp.sTipoPernocta = c.sIdPernocta ) '+

                      'left join moerecursos mr '+
                        'on ( mr.sIdRecurso = bp.sIdPersonal '+
                          'and mr.eTipoRecurso = "Personal" '+
                          'and mr.iIdMoe = ( select m.iIdMoe from moe m where m.sContrato = :contrato '+
                                            'and m.dIdFecha <= :fecha order by m.dIdFecha desc limit 1) ) '+

                      'left join personal p '+
                        'on ( p.sContrato = :contratoBarco '+
                          'and p.sIdPersonal = bp.sIdPersonal ) '+

                      'left join tiposdepersonal tp '+
                        'on ( p.sIdTipoPersonal = tp.sIdTipoPersonal ) '+

                      'left join bitacoradeactividades ba '+
                      'on (bp.sContrato =ba.sContrato and ba.sNumeroOrden=bp.sNumeroOrden and '+
                      'bp.dIdFecha = ba.didfecha and ba.iIdDiario=bp.iIdDiario and '+
                      'ba.iIdTarea=bp.iIdTarea and ba.iIdActividad=bp.iIdActividad ' +
                      'and ba.sIdTipoMovimiento = "E" ) '+

                      'left join ordenesdetrabajo ot '+
                        'on ( ot.sContrato = :contrato and ot.sNumeroOrden = ba.sNumeroOrden ) '+

                      'left join pernoctan pr '+
                        'on ( pr.sIdPernocta = ot.sIdPernocta ) '+

                      'left join plataformas pl '+
                        'on ( pl.sIdPlataforma = ot.sIdPlataforma ) '+
                      'where bp.snumeroorden=:Folio and c.sidpernocta=:Pernocta ' +
                      'and p.lpernocta="Si" ' +
                      //'where bp.sContrato = :contrato '+
                      //'and bp.dIdFecha = :fecha '+

                      'group by c.sIdPernocta '+
                      'order by c.sIdPernocta';
          parambyname('contrato').asstring := ReporteDiario.FieldByName('sOrden').asstring;
          parambyname('fecha').asdate := ReporteDiario.FieldByName('dIdFecha').asDatetime;
          parambyname('contratoBarco').asstring := global_contrato_barco;

        end;

        QrReportado.Active:=false;
        QrReportado.SQL.Text:='select c.*, a.sanexo,a.sdescripcion as anexo, a.stitulo as tituloAnexo '+
                              'from cuentas c '+
                              'left join anexos a on(a.sTipo= "PERNOCTA")';
        QrREportado.Open;


        QrAdicional.Active:=false;
        QrAdicional.SQL.Text:='select ifnull(sum(dCantidad),0) as dCantidad from bitacoradepernocta where ' +
                      'sContrato=:Contrato and dIdFecha=:fecha and sNumeroOrden=:Folio and '+
                      'sIdCuenta=:Pernocta ';

        DCiclo:=QrColumnas.RecordCount/TotalCol;
        iCiclo:=Trunc(DCiclo);
        if (dCiclo -iCiclo)>0 then
          Inc(iCiclo,1);

        iGrupo:=1;
        while iGrupo<=iCiclo do
        begin
          QrReportado.First;
          while not QrReportado.Eof do
          begin
            dTotal:=0;
            if iGrupo=1 then
              QrColumnas.First
            else
              QrColumnas.RecNo:=((iGrupo-1) * TotalCol)+ 1;

            mDatos.Append;
            mDatos.FieldByName('iGrupo').AsInteger:=Igrupo;
            mDatos.FieldByName('sidAnexo').AsString     := QrReportado.FieldByName('sanexo').asstring;
            mDatos.FieldByName('sAnexo').AsString       := QrReportado.FieldByName('anexo').asstring;
            mDatos.FieldByName('sTitulo').AsString      := QrReportado.FieldByName('tituloAnexo').asstring;
            mDatos.FieldByName('sidrecurso').AsString   := QrReportado.FieldByName('sidpernocta').asstring;
            mDatos.FieldByName('sdescripcion').AsString := QrReportado.fieldbyname('sdescripcion').asstring;
            mDatos.FieldByName('smedida').AsString      := QrReportado.fieldbyname('smedida').asstring;
            mDatos.FieldByName('dcantSol').AsFloat      :=0;
             mDatos.FieldByName('sTipo').AsString       :='';

            CuentaCol:=1;
            ValTmp:=0;
            CantTmp:=0;
            iPosTmp:=0;
            while not (QrColumnas.Eof) and (QrColumnas.RecNo<=((iGrupo) * TotalCol)) do
            begin
              mDatos.FieldByName('sNumeroOrden'+ Inttostr(CuentaCol)).AsString:=qrColumnas.FieldbyName('snumeroorden').AsString;
              mDatos.FieldByName('sPlataforma'+ Inttostr(CuentaCol)).AsString:=QrColumnas.Fieldbyname('splataforma').Asstring;

              with QrPernoctas do
              begin
                Active:=false;
                parambyname('Folio').AsString:=qrColumnas.FieldbyName('snumeroorden').AsString;
                parambyname('Pernocta').AsString:=QrReportado.FieldByName('sIdPernocta').AsString;
                Open;

                if Recordcount=0 then
                begin
                  ///Aqui va
                  QrAdicional.Active:=false;
                  QrAdicional.ParamByName('Contrato').AsString:=ReporteDiario.FieldByName('sOrden').asstring;
                  QrAdicional.ParamByName('Folio').AsString:=qrColumnas.FieldbyName('snumeroorden').AsString;
                  QrAdicional.ParamByName('Fecha').Asdate:=ReporteDiario.FieldByName('dIdFecha').asDatetime;
                  QrAdicional.ParamByName('Pernocta').AsString:= QrReportado.FieldByName('sIdCuenta').AsString;
                  QrAdicional.Open;
                  if QrAdicional.RecordCount=1 then
                  begin
                    mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat:=
                    xRound(QrAdicional.FieldByName('dCantidad').value,2);
                    dTotal:=dTotal+ xRound(QrAdicional.FieldByName('dCantidad').Value,2);
                    ValTmp:= ValTmp + (QrAdicional.FieldByName('dCantidad').Value);
                    if CantTmp<mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat then
                    begin
                      CantTmp:=mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat;
                      iPosTmp:= CuentaCol;
                    end;
                  end

                end;


                first;
                while not eof do
                begin

                  QrAdicional.Active:=false;
                  QrAdicional.ParamByName('Contrato').AsString:=ReporteDiario.FieldByName('sOrden').asstring;
                  QrAdicional.ParamByName('Folio').AsString:=qrColumnas.FieldbyName('snumeroorden').AsString;
                  QrAdicional.ParamByName('Fecha').Asdate:=ReporteDiario.FieldByName('dIdFecha').asDatetime;
                  QrAdicional.ParamByName('Pernocta').AsString:= QrReportado.FieldByName('sIdCuenta').AsString;
                  QrAdicional.Open;

                  if QrAdicional.RecordCount=1 then
                  begin
                    mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat:=xRound(FieldByName('dCantidad').Value,2)+
                    QrAdicional.FieldByName('dCantidad').AsFloat;
                    dTotal:=dTotal+ xRound(FieldByName('dCantidad').Value,2) +QrAdicional.FieldByName('dCantidad').AsFloat;
                    ValTmp:= ValTmp + (FieldByName('dCantidad').Value + QrAdicional.FieldByName('dCantidad').AsFloat);
                    if CantTmp < mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat then
                    begin
                      CantTmp := mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat;
                      iPosTmp:=CuentaCol;
                    end;
                  end
                  else
                  begin
                    mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat:=xRound(FieldByName('dCantidad').Value,2);
                    dTotal:=dTotal+ xRound(FieldByName('dCantidad').Value,2);
                    ValTmp:= ValTmp + FieldByName('dCantidad').Value;
                    if CantTmp < mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat then
                    begin
                      CantTmp := mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat;
                      iPosTmp:=CuentaCol;
                    end;
                  end;
                  next;
                end;
                //next;
              end;
              Inc(CuentaCol);
              QrColumnas.Next;
            end;
            if dTotal<>xRound(ValTmp,2) then
            begin
              mDatos.FieldByName('dcantidad' + Inttostr(iPosTmp)).AsFloat:=mDatos.FieldByName('dcantidad' + Inttostr(iPosTmp)).AsFloat
              + (xRound(ValTmp,2)-dTotal);
              mDatos.FieldByName('dcantTotal').AsFloat:=xRound(ValTmp,2);
            end
            else
              mDatos.FieldByName('dcantTotal').AsFloat:=dTotal;
            mDatos.Post;
            QrReportado.next;
          end;
          Inc(iGrupo,1);
        end;
        {$ENDREGION}

        {$REGION 'PARTIDAS DE PU'}

         {$REGION 'PARTIDAS PU.'}
        {Equipos...}
        QrRecursos.Active := False;
        QrRecursos.SQL.Clear;
        QrRecursos.SQL.Add( 'select b.sNumeroActividad, sum(b.dCantidad) as dCantidad '+
                            'from bitacoradeactividades b '+
                            'where b.sContrato = :Contrato and b.sNumeroOrden = :Orden and b.dIdFecha = :Fecha '+
                            'and b.sIdTipoMovimiento = "E" and b.sWbs = :Wbs group by b.sNumeroActividad');

        {$REGION 'CONSULTAS - PARTIDAS'}
        QrColumnas.Active:=false;
        QrColumnas.SQL.Clear;
        QrColumnas.SQL.Add( 'select ot.sContrato,ot.sIdFolio, '+
                                     'ot.sNumeroOrden, '+
                                     'ot.sIdPlataforma AS idPlataforma, '+
                                     'p.sDescripcion as Pernocta, '+
                                     'pt.sDescripcion as Plataforma '+
                              'from ordenesdetrabajo ot '+
                              'inner join contratos c '+
                              'on ( c.sContrato = ot.sContrato ) '+
                              'inner join bitacoradeactividades ba '+
                              'on ( ba.sContrato = ot.sContrato and ba.sNumeroOrden = ot.sNumeroOrden ) '+
                              'inner join pernoctan p '+
                              'on ( ot.sIdPernocta = p.sIdPernocta) '+
                              'inner join plataformas pt '+
                              'on ( ot.sIdPlataforma = pt.sIdPlataforma ) '+
                              'where c.sContrato = :contrato '+
                              'and ba.dIdFecha = :fecha '+
                              'group by ot.sIdFolio, p.sIdPernocta' );
        QrColumnas.ParamByName('Contrato').AsString:= ReporteDiario.FieldByName('sOrden').AsString;
        QrColumnas.ParamByName('Fecha').AsDate:=reportediario.FieldByName('dIdFecha').AsDateTime;
        QrColumnas.Open;
        {$ENDREGION}

        QrMoe.Active := False;
        QrMoe.SQL.Clear;
        QrMoe.SQL.Add('select b.sNumeroActividad as sIdRecurso, b.mDescripcion, a.sMedida, a.dCantidad, a.sAnexo, b.sWbs, '+
                      'n.sdescripcion as anexo, n.stitulo as tituloAnexo, n.sTipo '+
                      'from bitacoradeactividades b '+
                      'inner join actividadesxorden a on (a.sContrato = b.sContrato and a.sIdConvenio = :Convenio and a.sNumeroOrden  = a.sNumeroOrden '+
                      'and a.sTipoAnexo= "PU" and a.sNumeroActividad = b.sNumeroActividad and a.sWbs = b.sWbs and a.sTipoActividad = "Actividad") '+
                      'inner join anexos n on (a.sAnexo = n.sAnexo) '+
                      'where b.sContrato = :Contrato and b.dIdFecha = :Fecha '+
                      'and b.sIdTipoMovimiento = "E" group by n.iOrden, b.sContrato, b.sNumeroOrden, a.sNumeroActividad order by a.iItemOrden ');
        QrMoe.ParamByName('convenio').AsString := global_convenio;
        QrMoe.ParamByName('contrato').AsString := ReporteDiario.FieldByName('sOrden').AsString;
        QrMoe.ParamByName('Fecha').AsDateTime  := reportediario.FieldByName('dIdFecha').AsDateTime;
        QrMoe.Open;


        {$REGION 'INSERCION DE DATOS - INFORMACIÓN DEL PU'}

        DCiclo:=QrColumnas.RecordCount/TotalCol;
        iCiclo:=Trunc(DCiclo);
        if (dCiclo -iCiclo)>0 then
            Inc(iCiclo,1);
        iGrupo:=1;
        while iGrupo<=iCiclo do
        begin
            with qrMoe do
            begin
                first;
                while not eof do
                begin
                    mDatos.Append;
                    mDatos.FieldByName('iGrupo').AsInteger:=Igrupo;
                    mDatos.FieldByName('sidAnexo').AsString     :=FieldByName('sanexo').asstring;
                    mDatos.FieldByName('sAnexo').AsString       :=FieldByName('anexo').asstring;
                    mDatos.FieldByName('sTitulo').AsString      :=FieldByName('tituloAnexo').asstring;
                    mDatos.FieldByName('sidrecurso').AsString   :=FieldByName('sidrecurso').asstring;
                    mDatos.FieldByName('sdescripcion').AsString :=fieldbyname('mdescripcion').asstring;
                    mDatos.FieldByName('smedida').AsString      :=fieldbyname('smedida').asstring;
                    mDatos.FieldByName('dcantSol').AsFloat      :=fieldbyname('dcantidad').AsFloat;
                    mDatos.FieldByName('sTipo').AsString        :=FieldByName('sTipo').asstring;
                    dTotal:=0;
                    if iGrupo=1 then
                       QrColumnas.First
                    else
                       QrColumnas.RecNo:=((iGrupo-1) * TotalCol)+ 1;

                    CuentaCol:=1;
                    ValTmp:=0;
                    CantTmp:=0;
                    iPosTmp:=0;
                    while not (QrColumnas.Eof) and (QrColumnas.RecNo<=((iGrupo) * TotalCol)) do
                    begin
                        mDatos.FieldByName('sNumeroOrden'+ Inttostr(CuentaCol)).AsString := qrColumnas.FieldbyName('sNumeroOrden').AsString;
                        mDatos.FieldByName('sPernocta' + Inttostr(CuentaCol)).AsString   := QrColumnas.Fieldbyname('Pernocta').asstring;
                        mDatos.FieldByName('sPlataforma'+ Inttostr(CuentaCol)).AsString  := QrColumnas.Fieldbyname('Plataforma').Asstring;

                        QrRecursos.Active := False;
                        QrRecursos.ParamByName('Wbs').AsString      := FieldByName('sWbs').AsString;
                        QrRecursos.ParamByName('contrato').AsString := ReporteDiario.FieldByName('sOrden').AsString;
                        QrRecursos.ParamByName('Orden').AsString    := QrColumnas.FieldByName('sNumeroOrden').AsString;
                        QrRecursos.ParamByName('fecha').AsDate      := reportediario.FieldByName('dIdFecha').AsDateTime;
                        QrRecursos.Open;
                        //ABBY
                        if QrRecursos.RecordCount>0 then
                        begin
                          mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat:=xRound(QrRecursos.FieldByName('dCantidad').Value,4);
                          dTotal:=dTotal+ xRound(QrRecursos.FieldByName('dCantidad').Value,4);
                          ValTmp:=ValTmp + QrRecursos.FieldByName('dCantidad').Value;
                          if CantTmp<mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat then
                          begin
                            CantTmp:=mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat;
                            iPosTmp:=CuentaCol;
                          end;
                        end;
                        QrColumnas.next;
                        Inc(CuentaCol);
                    end;

                    if dTotal>=0 then
                    begin
                      if dTotal<>xRound(ValTmp,4) then
                      begin
                        mDatos.FieldByName('dcantidad' + Inttostr(iPosTmp)).AsFloat:=mDatos.FieldByName('dcantidad' + Inttostr(iPosTmp)).AsFloat
                        + (xRound(ValTmp,4)-dTotal);
                        mDatos.FieldByName('dcantTotal').AsFloat:=xRound(ValTmp,4);
                      end
                      else
                        mDatos.FieldByName('dcantTotal').AsFloat:=dTotal;
                      mDatos.Post;
                    end
                    else
                       mDatos.cancel;
                    next;
                end;
            end;
            Inc(iGrupo,1);
        end;
        {$ENDREGION}

    end;
    Td_Distribucion_detalle.DataSet:=MDatos;
    Td_Distribucion_detalle.FieldAliases.Clear;
    Td_ImpDistribucion_detalle.DataSet:=MImprime;
    Td_ImpDistribucion_detalle.FieldAliases.Clear;

    Reporte.DataSets.Add(Td_Distribucion_detalle);
    Reporte.DataSets.Add(Td_ImpDistribucion_detalle);
  finally
    QrRecursos.destroy;
    QrColumnas.destroy;
    QrMoe.destroy;
  end;
end;

Procedure TdConfiguracion(ParamContrato,ParamFolio:String;Var Reporte: TfrxReport);
var
  Td_contrato,
  Td_configuracion,
  Td_ConfiguracionOrden,
  TD_ConfigOTBarco: TfrxDBDataset;
  zContrato,
  zConfiguracion,
  QrConfigFolio,
  QrConfigBarco:TZReadOnlyQuery;
begin
  try
    try
      Td_contrato:= TfrxDBDataset.Create(nil);
      Td_contrato.UserName:='contrato';

      Td_configuracion:= TfrxDBDataset.Create(nil);
      Td_configuracion.UserName:='dsConfiguracion';

      Td_ConfiguracionOrden:= TfrxDBDataset.Create(nil);
      Td_ConfiguracionOrden.UserName:='DsConfiguracionOrden';

      TD_ConfigOTBarco:= TfrxDBDataset.Create(nil);
      TD_ConfigOTBarco.UserName:='TD_ConfigOTBarco';

      {Información del contrato}
      zContrato := TZReadOnlyQuery.Create(nil);
      zContrato.Connection := Connection.zConnection;
      zContrato.SQL.Add(' Select c.sLicitacion, c.sTitulo, c.sContrato, c.sTipoObra, c.mDescripcion, c.mCliente, c.bImagen, c.sUbicacion,  c.sCodigo, c.sLabelContrato, c.sCliente, '+
                'c.sIdResidencia, c.eLugarOT, '+
                'c2.sDescripcion as sConvenio, c2.dFechaInicio, c2.dFechaFinal, c2.dMontoMN, c2.dMontoDLL, '+
                'c2.dFecha, a.bImagen as bImagenActivo, a.sDescripcion as sDescripcionActivo, c.mComentarios '+
                'FROM contratos c '+
                'inner join activos a on (c.sIdActivo = a.sIdActivo) '+
                'inner join residencias rs on (c.sIdResidencia = rs.sIdResidencia) '+
                'inner join configuracion c3 on (c.sContrato = c3.sContrato) '+
                'inner join convenios c2 on (c3.sContrato = c2.sContrato And c3.sIdConvenio = c2.sIdConvenio) '+
                'Where c.sContrato = :Contrato ');
      zContrato.ParamByName('contrato').AsString := ParamContrato;
      zContrato.Open;

      {Información de la configuracion del sistema}
      zConfiguracion := TZReadOnlyQuery.Create(nil);
      zConfiguracion.Connection := Connection.zConnection;
      zConfiguracion.SQL.Add('select c.iFirmasReportes, c.sTipoPartida, c.sImprimePEP, ' +
                ' (select sContrato from contratos where sContrato =:contratobarco ) as sContratoBarco, ' +
                ' (select mDescripcion from contratos where sContrato =:contratobarco ) as mDescripcionBarco, ' +
                ' (select mcliente from contratos where sContrato =:contratobarco ) as mClienteBarco, ' +
                ' c.lLicencia, c.sReportesCIA, c.sLeyenda1, c.sLeyenda2, c.sLeyenda3, c.iFirmasGeneradores, ' +
                ' c.bImagen, c.sContrato, c.sNombre, c2.sCodigo, c.sPiePagina, c.sEmail, c.sWeb, c.sSlogan, c.sFirmasElectronicas, ' +
                ' c2.mDescripcion, c2.sTitulo, c2.mCliente, c2.bImagen as bImagenPEP, cv.dFechaInicio, cv.dfechaFinal ' +
                ',concat(c.sDireccion1," ",c.sDireccion2) as direccion,c.sCiudad,c.sTelefono,c.sFax'   +
                ' From contratos c2 '+
                ' INNER JOIN configuracion c ON (c.sContrato = c2.sContrato) ' +
                ' INNER JOIN convenios cv on (cv.sContrato = c2.sContrato and cv.sIdConvenio =:convenio) '+
                ' Where c2.sContrato = :Contrato');
      zConfiguracion.Params.ParamByName('contrato').DataType := ftString;
      zConfiguracion.Params.ParamByName('contrato').Value    := ParamContrato;
      zConfiguracion.Params.ParamByName('contratobarco').DataType := ftString;
      zConfiguracion.Params.ParamByName('contratobarco').Value    := global_contrato_barco;
      zConfiguracion.Params.ParamByName('convenio').DataType := ftString;
      zConfiguracion.Params.ParamByName('convenio').Value    := global_convenio;
      zConfiguracion.Open;

      QrConfigFolio := TZReadOnlyQuery.Create(nil);
      QrConfigFolio.Connection:=Connection.zConnection;
      QrConfigFolio.sql.Text:='select ot.*,p.sDescripcion as Plataforma from ordenesdetrabajo ot inner join plataformas p ' +
                              'on(p.sIdPlataforma=ot.sIdPlataforma) '+
                              'where ot.sContrato=:Contrato and ot.sNumeroOrden=:Orden';
      QrConfigFolio.ParamByName('Contrato').AsString:=ParamContrato;
      QrConfigFolio.ParamByName('Orden').AsString:=ParamFolio;
      try
        QrConfigFolio.Open;
      except
        raise;
      end;

      QrConfigBarco:=TzREadOnlyQuery.Create(nil);
      QrConfigBarco.Connection:=Connection.zConnection;
      QrConfigBarco.SQL.Text:='select * from contratos c inner join convenios cv '+
                              'on(c.sContrato and cv.sContrato) '+
                              'where c.sContrato=:ContratoBarco and cv.sIdConvenio=:Convenio';
      QrConfigBarco.ParamByName('ContratoBarco').AsString:= global_contrato_barco;
      QrConfigBarco.ParamByName('Convenio').AsString:= global_convenio;
      QrConfigBarco.Open;

      TD_ConfigOTBarco.DataSet:=QrConfigBarco;
      TD_ConfigOTBarco.FieldAliases.Clear;
      Td_contrato.DataSet:= zContrato;
      Td_contrato.FieldAliases.Clear;
      Td_configuracion.DataSet:= zConfiguracion;
      Td_configuracion.FieldAliases.Clear;
      Td_ConfiguracionOrden.DataSet:= QrConfigFolio;
      Td_ConfiguracionOrden.FieldAliases.Clear;



      Reporte.DataSets.Add(Td_contrato);
      Reporte.DataSets.Add(Td_configuracion);
      Reporte.DataSets.Add(Td_ConfiguracionOrden);
      Reporte.DataSets.Add(TD_ConfigOTBarco);

    except

    end;

  finally

  end;


end;

Procedure FirmasPDF_Generales(Reportediario :TzREadOnlyQuery;Var Reporte: TfrxReport;Tipo:FtTipo);
Var
  zFirmas : TZReadOnlyQuery;
begin
  Try
    {Variables globales para Firmantes..}
    sSuperIntendente      := 'SIN FIRMANTE';
    sSuperIntendentePatio := 'SIN FIRMANTE';
    sRepresentanteTecnico := 'SIN FIRMANTE';
    sSupervisor           := 'SIN FIRMANTE';
    sSupervisorPatio      := 'SIN FIRMANTE';
    sSupervisorGenerador  := 'SIN FIRMANTE';
    sSupervisorEstimacion := 'SIN FIRMANTE';
    sSupervisorSubContratista   := 'SIN FIRMANTE';
    sResidente                  := 'SIN FIRMANTE';
    sPuestoSupervisorSubContratista := 'SIN PUESTO';
    sPuestoSuperintendente      := 'SIN PUESTO';
    sPuestoSupervisor           := 'SIN PUESTO';
    sPuestoSupervisorGenerador  := 'SIN PUESTO';
    sPuestoSupervisorEstimacion := 'SIN PUESTO';
    sSupervisorTierra           := 'SIN PUESTO';
    sPuestoSupervisorTierra     := 'SIN PUESTO';
    sPuestoRepresentanteTecnico := 'SIN PUESTO';
    sPuestoResidente            := 'SIN PUESTO';

    zFirmas := tzReadOnlyQuery.Create(nil);
    zFirmas.Connection := connection.zconnection;
    zFirmas.Active := False;
    zFirmas.SQL.Clear;
    if Reportediario.FieldByName('sNumeroOrden').AsString <> '' then
    begin
      zFirmas.SQL.Add('Select * from firmas where sContrato = :contrato and sIdTurno =:Turno and sNumeroOrden = :Orden And dIdFecha = :fecha');
      zFirmas.Params.ParamByName('Orden').DataType := ftString;
      zFirmas.Params.ParamByName('Orden').Value    := Reportediario.FieldByName('sNumeroOrden').AsString;
    end
    else
      zFirmas.SQL.Add('Select * from firmas where sContrato = :contrato and sIdTurno =:Turno And dIdFecha = :fecha');
    zFirmas.Params.ParamByName('Contrato').DataType := ftString;
    zFirmas.Params.ParamByName('Contrato').Value    := Reportediario.FieldByName('sOrden').AsString;
    zFirmas.Params.ParamByName('Turno').DataType    := ftString;
    zFirmas.Params.ParamByName('Turno').Value       := Reportediario.FieldByName('sIdTurno').AsString;
    zFirmas.Params.ParamByName('fecha').DataType    := ftDate;
    zFirmas.Params.ParamByName('fecha').Value       := Reportediario.FieldByName('dIdFecha').AsDateTime;
    zFirmas.Open;

    if zFirmas.RecordCount > 0 then
    begin
      sSuperintendente      := zFirmas.FieldValues['sFirmante1'];
      sSupervisor           := zFirmas.FieldValues['sFirmante2'];
      sSupervisorGenerador  := zFirmas.FieldValues['sFirmante3'];
      sSupervisorEstimacion := zFirmas.FieldValues['sFirmante4'];
      sSupervisorTierra     := zFirmas.FieldValues['sFirmante5'];
      sResidente            := zFirmas.FieldValues['sFirmante6'];
      sSuperintendentePatio := zFirmas.FieldValues['sFirmante7'];
      sSupervisorPatio      := zFirmas.FieldValues['sFirmante8'];
      sRepresentanteTecnico := zFirmas.FieldValues['sFirmante9'];
      sSupervisorSubContratista := zFirmas.FieldValues['sFirmante10'];

      sPuestoSuperintendente      := zFirmas.FieldValues['sPuesto1'];
      sPuestoSupervisor           := zFirmas.FieldValues['sPuesto2'];
      sPuestoSupervisorGenerador  := zFirmas.FieldValues['sPuesto3'];
      sPuestoSupervisorEstimacion := zFirmas.FieldValues['sPuesto4'];
      sPuestoSupervisorTierra     := zFirmas.FieldValues['sPuesto5'];
      sPuestoResidente            := zFirmas.FieldValues['sPuesto6'];
      sPuestoSuperintendentePatio := zFirmas.FieldValues['sPuesto7'];
      sPuestoSupervisorPatio      := zFirmas.FieldValues['sPuesto8'];
      sPuestoRepresentanteTecnico := zFirmas.FieldValues['sPuesto9'];
      sPuestoSupervisorSubContratista := zFirmas.FieldValues['sPuesto10'];
    end
    else
    begin
      zFirmas.Active := False;
      zFirmas.SQL.Clear;
      if Reportediario.FieldByName('sNumeroOrden').AsString <> '' then
      begin
        zFirmas.SQL.Add('Select * from firmas where sContrato = :contrato and sNumeroOrden = :Orden and sIdTurno =:Turno And dIdFecha <= :fecha Order By dIdFecha DESC');
        zFirmas.Params.ParamByName('Orden').DataType := ftString;
        zFirmas.Params.ParamByName('Orden').Value    := Reportediario.FieldByName('sNumeroOrden').AsString;
      end
      else
        zFirmas.SQL.Add('Select * from firmas where sContrato = :contrato and sIdTurno =:Turno And dIdFecha <= :fecha Order By dIdFecha DESC');

      zFirmas.Params.ParamByName('Contrato').DataType := ftString;
      zFirmas.Params.ParamByName('Contrato').Value    := Reportediario.FieldByName('sOrden').AsString;
      zFirmas.Params.ParamByName('Turno').DataType    := ftString;
      zFirmas.Params.ParamByName('Turno').Value       := Reportediario.FieldByName('sIdTurno').AsString;
      zFirmas.Params.ParamByName('fecha').DataType    := ftDate;
      zFirmas.Params.ParamByName('fecha').Value       := Reportediario.FieldByName('dIdFecha').AsDateTime;
      zFirmas.Open;

      if zFirmas.RecordCount > 0 then
      begin
        sSuperintendente      := zFirmas.FieldValues['sFirmante1'];
        sSupervisor           := zFirmas.FieldValues['sFirmante2'];
        sSupervisorGenerador  := zFirmas.FieldValues['sFirmante3'];
        sSupervisorEstimacion := zFirmas.FieldValues['sFirmante4'];
        sSupervisorTierra     := zFirmas.FieldValues['sFirmante5'];
        sResidente            := zFirmas.FieldValues['sFirmante6'];
        sSuperintendentePatio := zFirmas.FieldValues['sFirmante7'];
        sSupervisorPatio      := zFirmas.FieldValues['sFirmante8'];
        sRepresentanteTecnico := zFirmas.FieldValues['sFirmante9'];
        sSupervisorSubContratista := zFirmas.FieldValues['sFirmante10'];

        sPuestoSuperintendente      := zFirmas.FieldValues['sPuesto1'];
        sPuestoSupervisor           := zFirmas.FieldValues['sPuesto2'];
        sPuestoSupervisorGenerador  := zFirmas.FieldValues['sPuesto3'];
        sPuestoSupervisorEstimacion := zFirmas.FieldValues['sPuesto4'];
        sPuestoSupervisorTierra     := zFirmas.FieldValues['sPuesto5'];
        sPuestoResidente            := zFirmas.FieldValues['sPuesto6'];
        sPuestoSuperintendentePatio := zFirmas.FieldValues['sPuesto7'];
        sPuestoSupervisorPatio      := zFirmas.FieldValues['sPuesto8'];
        sPuestoRepresentanteTecnico := zFirmas.FieldValues['sPuesto9'];
        sPuestoSupervisorSubContratista := zFirmas.FieldValues['sPuesto10'];
      end
    end;
    zFirmas.Destroy;

  Finally

  End;
end;

Procedure EncabezadoPDF_Horizontal(Reportediario :TzREadOnlyQuery;Var Reporte: TfrxReport;Tipo:FtTipo);
const
   Dias: array[1..7] of string = ('LUNES', 'MARTES', 'MIERCOLES', 'JUEVES', 'VIERNES', 'SABADO', 'DOMINGO');
Var
  zContrato, zConfiguracion,
  zEmbarcacion, zDuracion     : TZReadOnlyQuery;
  Td_contrato, Td_configuracion,
  Td_embarcacion, Td_duracion : TfrxDBDataset;
  iDia : integer;
  sDia : string;

begin
  Try

    Td_contrato:= TfrxDBDataset.Create(nil);
    Td_contrato.UserName:='contrato';

    Td_configuracion:= TfrxDBDataset.Create(nil);
    Td_configuracion.UserName:='dsConfiguracion';

    Td_embarcacion:= TfrxDBDataset.Create(nil);
    Td_embarcacion.UserName:='dsEmbarcacion';

    Td_duracion:= TfrxDBDataset.Create(nil);
    Td_duracion.UserName:='ds_Duracion';

    {Información del contrato}
    zContrato := TZReadOnlyQuery.Create(nil);
    zContrato.Connection := Connection.zConnection;
    zContrato.SQL.Add(' Select c.sLicitacion, c.sTitulo, c.sContrato, c.sTipoObra, c.mDescripcion, c.mCliente, c.bImagen, c.sUbicacion,  c.sCodigo, c.sLabelContrato, c.sCliente, '+
              'c.sIdResidencia, c.eLugarOT, '+
              'c2.sDescripcion as sConvenio, c2.dFechaInicio, c2.dFechaFinal, c2.dMontoMN, c2.dMontoDLL, '+
              'c2.dFecha, a.bImagen as bImagenActivo, a.sDescripcion as sDescripcionActivo, c.mComentarios '+
              'FROM contratos c '+
              'inner join activos a on (c.sIdActivo = a.sIdActivo) '+
              'inner join residencias rs on (c.sIdResidencia = rs.sIdResidencia) '+
              'inner join configuracion c3 on (c.sContrato = c3.sContrato) '+
              'inner join convenios c2 on (c3.sContrato = c2.sContrato And c3.sIdConvenio = c2.sIdConvenio) '+
              'Where c.sContrato = :Contrato ');
    zContrato.ParamByName('contrato').AsString := reportediario.FieldByName('sOrden').AsString;
    zContrato.Open;

    {Información de la configuracion del sistema}
    zConfiguracion := TZReadOnlyQuery.Create(nil);
    zConfiguracion.Connection := Connection.zConnection;
    zConfiguracion.SQL.Add('select c.iFirmasReportes, c.sTipoPartida, c.sImprimePEP, ' +
              ' (select sContrato from contratos where sContrato =:contratobarco ) as sContratoBarco, ' +
              ' (select mDescripcion from contratos where sContrato =:contratobarco ) as mDescripcionBarco, ' +
              ' c.lLicencia, c.sReportesCIA, c.sLeyenda1, c.sLeyenda2, c.sLeyenda3, c.iFirmasReportes, ' +
              ' c.bImagen, c.sContrato, c.sNombre, c2.sCodigo, c.sPiePagina, c.sEmail, c.sWeb, c.sSlogan, c.sFirmasElectronicas, ' +
              ' c2.mDescripcion, c2.sTitulo, c2.mCliente, c2.bImagen as bImagenPEP, cv.dFechaInicio, cv.dfechaFinal ' +
              ' From contratos c2 '+
              ' INNER JOIN configuracion c ON (c.sContrato = c2.sContrato) ' +
              ' INNER JOIN convenios cv on (cv.sContrato = c2.sContrato and cv.sIdConvenio =:convenio) '+
              ' Where c2.sContrato = :Contrato');
    zConfiguracion.Params.ParamByName('contrato').DataType := ftString;
    zConfiguracion.Params.ParamByName('contrato').Value    := reportediario.FieldByName('sOrden').AsString;
    zConfiguracion.Params.ParamByName('contratobarco').DataType := ftString;
    zConfiguracion.Params.ParamByName('contratobarco').Value    := global_contrato_barco;
    zConfiguracion.Params.ParamByName('convenio').DataType := ftString;
    zConfiguracion.Params.ParamByName('convenio').Value    := global_convenio;
    zConfiguracion.Open;

    {Busqueda de la embarcacion}
    zEmbarcacion := TZReadOnlyQuery.Create(nil);
    zEmbarcacion.Connection := Connection.zConnection;
    zEmbarcacion.SQL.Add('SELECT em.sDescripcion,em.sContrato, em.sIdEmbarcacion, cc.sLocalizacion, cc.sCantidad as CantidadOlas,'+
                  'd.sDescripcion as DireccionOlas, '+
                  '(select sCantidad from condicionesclimatologicas where sContrato =:contrato and dIdFecha = :fecha and sIdTurno = :turno and iIdCondicion = 1 group by iIdCondicion ) as CantidadViento, '+
                  '(select d2.sDescripcion from condicionesclimatologicas cc2 '+
                  'inner join direcciones d2 on (d2.iIdDireccion = cc2.iIdDireccion)'+
                  'where cc2.sContrato =:contrato and cc2.dIdFecha = :fecha and cc2.sIdTurno =:turno and cc2.iIdCondicion = 1 group by iIdCondicion ) as DireccionViento '+
                  'FROM embarcacion_vigencia AS ev '+
                  'INNER JOIN embarcaciones AS em ON (em.sIdEmbarcacion = ev.sIdEmbarcacion) '+
                  'left join condicionesclimatologicas as cc on (cc.sContrato = ev.sContrato and cc.dIdFecha =:fecha and cc.sIdTurno =:turno and cc.iIdCondicion = 2) '+
                  'left join direcciones as d on (d.iIdDireccion = cc.iIdDireccion) '+
                  'WHERE ev.sContrato = :contrato and ev.dFechaInicio = (Select max(ev2.dfechainicio) '+
                  'from embarcacion_vigencia ev2 where ev.sContrato = em.sContrato and ev2.dfechainicio <= :fecha) order by sHorario DESC ');
    zEmbarcacion.ParamByName('Contrato').AsString := Global_Contrato_Barco;
    zEmbarcacion.ParamByName('fecha').AsDateTime  := reportediario.FieldByName('dIdfecha').AsDateTime;
    zEmbarcacion.ParamByName('turno').AsString    := reportediario.FieldByName('sIdTurno').AsString;
    zEmbarcacion.Open;

    {Este bloque es para obtener el día}
     iDia := DayOfTheWeek(Reportediario.FieldByName('dIdFecha').AsDateTime);
     sDia := dias[iDia];

    {Consulto los días de contrato y de vigencia}
    zDuracion := TZReadOnlyQuery.Create(nil);
    zDuracion.Connection := Connection.zConnection;
    zDuracion.SQL.Add('SELECT ' +
                '	MIN(dFechaInicio) AS dInicioDeContrato, ' +
                '	MAX(dFechaFinal) AS dFinalDeContrato,   ' +
                '	DATEDIFF(MAX(dFechaFinal),MIN(dFechaInicio)) + 1 AS dDiasDeContrato, ' +
                '	DATEDIFF(MAX(dFechaFinal), :Hoy) AS dDiasRestantes,     ' +
                '	DATEDIFF(:Hoy, MIN(dFechaInicio)) + 1 AS dDiasTranscurridos, '+
                ' DATE_FORMAT(:hoy,"%d/%m/%Y") as dIdFecha, '+
                ' :Dia as DiaSemana '+
                ' FROM convenios WHERE sContrato = :Orden');
    zDuracion.ParamByName('Hoy').AsDate      := Reportediario.FieldByName('dIdFecha').AsDateTime;
    zDuracion.ParamByName('Orden').AsString  := ReporteDiario.FieldByName('sOrden').AsString;
    zDuracion.ParamByName('Dia').AsString    := sDia;
    zDuracion.Open;

    Td_contrato.DataSet:= zContrato;
    Td_contrato.FieldAliases.Clear;
    Td_configuracion.DataSet:= zConfiguracion;
    Td_configuracion.FieldAliases.Clear;
    Td_embarcacion.DataSet:= zEmbarcacion;
    Td_embarcacion.FieldAliases.Clear;
    Td_duracion.DataSet:= zDuracion;
    Td_duracion.FieldAliases.Clear;

    Reporte.DataSets.Add(Td_contrato);
    Reporte.DataSets.Add(Td_configuracion);
    Reporte.DataSets.Add(Td_embarcacion);
    Reporte.DataSets.Add(Td_duracion);

  Finally

  End;
end;


end.
