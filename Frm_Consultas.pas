unit Frm_Consultas;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, cxGraphics, cxControls, cxLookAndFeels, cxLookAndFeelPainters,
  cxContainer, cxEdit, dxSkinsCore, dxSkinDevExpressStyle, dxSkinFoggy,
  cxGroupBox, cxStyles, dxSkinscxPCPainter, cxGrid, dxLayoutContainer,
  dxLayoutControl, ComCtrls, dxCore, cxDateUtils, dxLayoutcxEditAdapters,
  cxTextEdit, cxMaskEdit, cxDropDownEdit, cxCalendar, dxCheckGroupBox,
  cxRadioGroup, cxProgressBar, cxLookupEdit, cxDBLookupEdit, cxDBLookupComboBox,
  cxCustomData, cxFilter, cxData, cxDataStorage, cxNavigator, DB, cxDBData,
  cxMemo, cxGridCustomTableView, cxGridTableView, cxGridBandedTableView,
  cxGridDBBandedTableView, cxGridCustomView, cxClasses, cxGridLevel, dxmdaset,
  ZAbstractRODataset, ZDataset, StdCtrls, Mask, JvExMask, JvToolEdit, JvCombobox,
  cxSSheet, dxLayoutControlAdapters, Menus, cxButtons, UnitTarifa, dxSkinBlack,
  dxSkinBlue, dxSkinBlueprint, dxSkinCaramel, dxSkinCoffee, dxSkinDarkRoom,
  dxSkinDarkSide, dxSkinDevExpressDarkStyle, dxSkinGlassOceans,
  dxSkinHighContrast, dxSkiniMaginary, dxSkinLilian, dxSkinLiquidSky,
  dxSkinLondonLiquidSky, dxSkinMcSkin, dxSkinMetropolis, dxSkinMetropolisDark,
  dxSkinMoneyTwins, dxSkinOffice2007Black, dxSkinOffice2007Blue,
  dxSkinOffice2007Green, dxSkinOffice2007Pink, dxSkinOffice2007Silver,
  dxSkinOffice2010Black, dxSkinOffice2010Blue, dxSkinOffice2010Silver,
  dxSkinOffice2013DarkGray, dxSkinOffice2013LightGray, dxSkinOffice2013White,
  dxSkinPumpkin, dxSkinSeven, dxSkinSevenClassic, dxSkinSharp, dxSkinSharpPlus,
  dxSkinSilver, dxSkinSpringTime, dxSkinStardust, dxSkinSummer2008,
  dxSkinTheAsphaltWorld, dxSkinsDefaultPainters, dxSkinValentine, dxSkinVS2010,
  dxSkinWhiteprint, dxSkinXmas2008Blue;

type
  FtConsulta=(FtDiario,FtAcumulado);

type
  TFrmConsultas = class(TForm)
    GBx1: TcxGroupBox;
    GBx2: TcxGroupBox;
    GBx3: TcxGroupBox;
    dxLayoutControl1Group_Root: TdxLayoutGroup;
    dxLayoutControl1: TdxLayoutControl;
    dxLayoutControl1Item1: TdxLayoutItem;
    DtEdtFechaFin: TcxDateEdit;
    dxLayoutControl1Item2: TdxLayoutItem;
    DtEdtFechaInicio: TcxDateEdit;
    ChkGbxFolios: TdxCheckGroupBox;
    dxLayoutControl1Item3: TdxLayoutItem;
    dxLayoutControl1Group1: TdxLayoutAutoCreatedGroup;
    RdGpTipo: TcxRadioGroup;
    dxLayoutControl1Item4: TdxLayoutItem;
    dxLayoutControl1Group2: TdxLayoutAutoCreatedGroup;
    ProgresoReporte: TcxProgressBar;
    dxLayoutControl1Item5: TdxLayoutItem;
    dxLayoutControl1Group3: TdxLayoutAutoCreatedGroup;
    dxLayoutControl2Group_Root: TdxLayoutGroup;
    dxLayoutControl2: TdxLayoutControl;
    ChkGbxPartidas: TdxCheckGroupBox;
    dxLayoutControl1Item6: TdxLayoutItem;
    dxLayoutControl1Group4: TdxLayoutAutoCreatedGroup;
    ChkCmbFolios: TJvCheckedComboBox;
    dxLayoutControl2Item1: TdxLayoutItem;
    SprShBkConsulta: TcxSpreadSheetBook;
    btnGuardar: TcxButton;
    dxLayoutControl1Item7: TdxLayoutItem;
    dxLayoutControl1Group5: TdxLayoutAutoCreatedGroup;
    dxLayoutControl3Group_Root: TdxLayoutGroup;
    dxLayoutControl3: TdxLayoutControl;
    ChkCmbPartidas: TJvCheckedComboBox;
    dxLayoutControl3Item1: TdxLayoutItem;
    btnConsultar: TcxButton;
    dxLayoutControl1Item8: TdxLayoutItem;
    dxLayoutControl1Group6: TdxLayoutAutoCreatedGroup;
    QrFolios: TZReadOnlyQuery;
    QrPartidas: TZReadOnlyQuery;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }
    procedure ConsultaAvances(FechaI,FechaT:TDate;Tipo:FtConsulta=FtDiario;Folios:string='');
    procedure ModificarEstructura(FechaI,FechaT:TDate;Tipo:FtConsulta);
   // procedure ConsultaExLoad(pTipo:RtRecurso;Libro:TcxSpreadSheetBook;pFechaInicio,pFechaTermino:TDate;pFolios,pPartidas:String);
  public
    { Public declarations }
  end;
  //RtRecurso=(RtAll,RtPersonal,RtEquipo,RtPernocta,RtBarco,RtExtraordinaria);
var
  FrmConsultas: TFrmConsultas;

implementation

uses frm_connection, global;

{$R *.dfm}

procedure TFrmConsultas.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Action:=caFree;
end;
(*
procedure TFrmConsultas.ConsultaExLoad(pTipo:RtRecurso;Libro:TcxSpreadSheetBook;pFechaInicio,pFechaTermino:TDate;pFolios,pPartidas:String);
const
  SQlRef: array[1..2,1..4] of string=(('bitacoradepersonal','personal','sIdPersonal','PERSONAL'),('bitacoradeequipos','equipos','sIdEquipo','EQUIPO'));

var
  ACellObj: TcxSSCellObject;
  IRow,pIRow:Integer;
  NPdas:Integer;
  Hoja:TcxssBookSheet;
  i:Byte;
  QrActividades,QrAvance:TZReadOnlyQuery;
  QImportes,
  QRecursos:TZQuery;
  TotalHrs,sFormulaSumMn,sFormulaSumDll:string;
begin

  QrActividades:=TZReadOnlyQuery.Create(nil);
  QrAvance:=TZReadOnlyQuery.Create(nil);
  QImportes:=TZQuery.Create(nil);
  QRecursos:=TZQuery.Create(nil);
  try
    QrActividades.Connection:=connection.zConnection;
    QImportes.Connection:=connection.zConnection;
    QRecursos.Connection:=connection.zConnection;
    QrAvance.Connection:=connection.zConnection;

    QrActividades.SQL.Text:='select ao.* from actividadesxorden ao inner join acta_campo ac '+
                            'on (ac.sContrato=ao.sContrato and ac.sNumeroOrden=ao.sNumeroOrden and ' +
                            'ac.swbs=ao.swbs and ac.sNumeroActividad=ao.sNumeroActividad) ' +
                            'where ac.iIdActa=:Acta and ao.sTipoActividad=:Tipo ' +
                            'group by ao.swbs order by ao.iItemOrden';

    QrAvance.Connection:=connection.zConnection;
    QrAvance.SQL.Text:= 'select b.*, ' +
                        '( SELECT (ifnull(sum(ba.dAvance), 0)) ' +
                                          '		FROM ' +
                                          '			bitacoradeactividades AS ba ' +
                                          '		WHERE ' +
                                          '			ba.sContrato = b.sContrato ' +
                                          '		AND ba.sNumeroOrden = b.sNumeroOrden ' +
                                          '		AND ba.sIdTipoMovimiento = b.sIdTipoMovimiento ' +
                                          '		AND ba.swbs = b.swbs ' +
                                          '		AND ba.sNumeroActividad = b.sNumeroActividad ' +
                                          '		AND ( ba.didfecha < b.didfecha OR (ba.didfecha = b.didfecha AND cast(ba.sHoraInicio AS Time) '+
                                          '   < cast(b.sHoraInicio AS Time))  )	) AS AvanceAnterior ' +
                        ' from bitacoradeactividades b' + #13#10 +
                        'where b.sContrato=:Contrato and b.snumeroorden=:Orden and b.sNumeroActividad=:Actividad' + #13#10 +
                        'and b.sIdTipoMovimiento=:Tipo' + #13#10 +
                        'order by b.didfecha,time(b.sHoraInicio)' ;

    QImportes.SQL.Text:='select * from acta_campo where iIdActa=:Acta and swbs=:wbs and sNumeroActividad=:Actividad and '   +
                        'sIdRecurso like "$IMPORTE%" order by iOrdenTipo';

    QRecursos.SQL.Text:='select * from acta_campo where iIdActa=:Acta and swbs=:wbs and sNumeroActividad=:Actividad and '   +
                        'eTipo=:Tipo and sIdRecurso not like "$IMPORTE%" order by iOrdenRecurso';


    { QrRecursos.Active:=False;
        QrRecursos.SQL.Text:= 'select br.' +SQlRef[i,3] + ' as sIdRecurso,br.sDescripcion,r.sMedida,sum(br.dCanthh) as dCanthh,sum(Ifnull(br.dAjuste,0)) as Ajuste,r.dVentaMn,r.dVentaDll ' +
                              'from '+ SQlRef[i,1]  + ' br ' +
                              'left join ' + SQlRef[i,2] + ' r ' +
                              'on(r.sContrato=:Contrato and br.'+SQlRef[i,3]+'=r.'+ SQlRef[i,3] +') ' +
                              'where br.sContrato=:Orden and br.sNumeroOrden=:Folio and br.sNumeroActividad=:Actividad ' +
                              'group by sIdRecurso,br.dIdFecha order by r.iitemorden' ;
        QrRecursos.ParamByName('contrato').AsString:=global_Contrato_Barco;
        QrRecursos.ParamByName('Orden').AsString:=DatosActa.FieldByName('sContrato').AsString;
        QrRecursos.ParamByName('Folio').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
        QrRecursos.ParamByName('Actividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
        QrRecursos.Open;
}





    for I := 3 downto 0 do
    begin
      Hoja:=Libro.Pages[i];

      if i=0 then
        Hoja.Caption:='ACTA DE ACTIVIDADES';

      if i=1 then
        Hoja.Caption:='COSTO POR ACTIVIDAD';

      if i=2 then
        Hoja.Caption:='CAMPO';

      if i=3 then
        Hoja.Caption:='DESGLOSE DE COSTOS';



      QrActividades.Active:=False;
      QrActividades.ParamByName('Acta').AsInteger:=Datos.FieldByName('iIdActa').AsInteger;
      if (I=2) or (I=1) then
        QrActividades.ParamByName('Tipo').AsString:='Actividad'
      else
        QrActividades.ParamByName('Tipo').AsString:='Paquete';
      QrActividades.Open;

      Hoja.ClearAll;
      //NPdas:=0;

      IRow:=0;
      while not QrActividades.Eof do
      begin
        sFormulaSumMn:='';
        sFormulaSumDll:='';
        QImportes.Active:=False;
        QImportes.ParamByName('Acta').AsInteger:=Datos.FieldByName('iIdActa').AsInteger;
        QImportes.ParamByName('wbs').AsString:=QrActividades.FieldByName('swbs').AsString;
        QImportes.ParamByName('Actividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
        QImportes.Open;

        with Hoja do
        begin
          //IRow:=NPdas;
          if (Hoja.Caption='CAMPO') or (Hoja.Caption='COSTO POR ACTIVIDAD') then
          begin
            Hoja.Rows.ResetDefault(IRow);
            Hoja.Rows.ResetDefault(IRow+1);
            Cols.Size[7]:=Cols.Size[7] - 5;

            ACellObj := GetCellObject(0, IRow);
            FormatoCelda(ACellObj,FrTitle);
            ACellObj.SetCellText('PARTIDA');
            ACellObj := GetCellObject(2, IRow);
            FormatoCelda(ACellObj,FrTitle);
            ACellObj.SetCellText('ACTIVIDAD');

            GetCellObject(2, IRow+1).Style.WordBreak:=True;

           //
            Rows.Size[IRow+1] := 70;

            LockUnLock(Hoja,0,1,IRow,0,False);
            RangeMerge(Hoja,0,1,IRow,0);
            LockUnLock(Hoja,0,1,IRow,0);

            LockUnLock(Hoja,2,10,IRow,0,False);
            RangeMerge(Hoja,2,10,IRow,0);
            LockUnLock(Hoja,2,10,IRow,0);

            LockUnLock(Hoja,0,1,IRow+1,0,False);
            RangeMerge(Hoja,0,1,IRow+1,0);
            LockUnLock(Hoja,0,1,IRow+1,0);

            LockUnLock(Hoja,2,10,IRow+1,0,False);
            RangeMerge(Hoja,2,10,IRow+1,0);
            LockUnLock(Hoja,2,10,IRow+1,0);

            BordeCelda(Hoja,0,10,IRow,1);
            Hoja.GetCellObject(0,IRow+1).SetCellText(QrActividades.FieldByName('sNumeroActividad').AsString);
            Hoja.GetCellObject(2,IRow+1).SetCellText(Trim(QrActividades.FieldByName('mDescripcion').AsString));
            AlineacionCelda(Hoja,haCENTER,vaCENTER,0,0,IRow+1,0);


            if (Hoja.Caption='CAMPO') then
            begin
              inc(IRow,3);
              Hoja.Rows.ResetDefault(IRow);
              ACellObj := GetCellObject(1, IRow);
              FormatoCelda(ACellObj,FrSubTitle);
              ACellObj.SetCellText('PERIODOS DE EJECUCION DE LA ACTIVIDAD');
              LockUnLock(Hoja,1,8,IRow,0,False);
              RangeMerge(Hoja,1,8,IRow,0);
              LockUnLock(Hoja,1,8,IRow,0);
              BordeCelda(Hoja,1,8,IRow,2);
                        //IRow:=5;
              Inc(IRow);
              Hoja.Rows.ResetDefault(IRow);
              ACellObj := GetCellObject(1, IRow);
              FormatoCelda(ACellObj,FrSubTitle);
              ACellObj.SetCellText('FECHA');
              ACellObj := GetCellObject(2,IRow);
              FormatoCelda(ACellObj,FrSubTitle);
              ACellObj.SetCellText('INICIO');
              ACellObj := GetCellObject(3, IRow);
              FormatoCelda(ACellObj,FrSubTitle);
              ACellObj.SetCellText('TERMINO');
              ACellObj := GetCellObject(4, IRow);
              FormatoCelda(ACellObj,FrSubTitle);
              ACellObj.SetCellText('AFECTACIÓN');
              ACellObj := GetCellObject(5, IRow);
              FormatoCelda(ACellObj,FrSubTitle);
              ACellObj.Style.WordBreak:=True;
              ACellObj.SetCellText('INTERVALO TIEMPO');
              ACellObj := GetCellObject(6, IRow);
              FormatoCelda(ACellObj,FrSubTitle);
              ACellObj.Style.WordBreak:=True;
              ACellObj.SetCellText('AVANCE ANTERIOR');
              ACellObj := GetCellObject(7, IRow);
              FormatoCelda(ACellObj,FrSubTitle);
              ACellObj.Style.WordBreak:=True;
              ACellObj.SetCellText('AVANCE ACTUAL');
              ACellObj := GetCellObject(8,IRow);
              FormatoCelda(ACellObj,FrSubTitle);
              ACellObj.Style.WordBreak:=True;
              ACellObj.SetCellText('AVANCE ACUMULADO');
              Rows.Size[IRow] :=30;

              inc(IRow,2);
              //IRow:=7;
              Hoja.Rows.ResetDefault(IRow);
              ACellObj := GetCellObject(1, IRow);
              FormatoCelda(ACellObj,FrTitle);
              ACellObj.SetCellText('DURACION TIEMPO EFECTIVO (HRS):');

                           //Inc(IRow,3);
              LockUnLock(Hoja,1,4,IRow,0,False);
              RangeMerge(Hoja,1,4,IRow,0);
              LockUnLock(Hoja,1,4,IRow,0);
              BordeCelda(Hoja,1,5,IRow,0);
              IRow:=IRow-1;

              QrAvance.Active:=False;
              QrAvance.ParamByName('contrato').AsString:=Datos.FieldByName('sContrato').AsString;
              QrAvance.ParamByName('Orden').AsString:=Datos.FieldByName('sNumeroOrden').AsString;
              QrAvance.ParamByName('Actividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
              QrAvance.ParamByName('tipo').AsString:='ED';
              QrAvance.Open;
              TotalHrs:='00:00';
              while not QrAvance.Eof do
              begin
                if QrAvance.RecordCount<>QrAvance.RecNo then
                begin
                  LockUnLock(Hoja,1,8,IRow,0,false);
                  Hoja.SelectCell(1,IRow);
                  Hoja.InsertCells(Hoja.SelectionRect,msallRow );
                end
                else
                  LockUnLock(Hoja,1,8,IRow,0,true);
                //Hoja.GetCellObject(1,Row).SetCellText(QrAvance.FieldByName('dIdFecha').AsString);
                Hoja.GetCellObject(1,IRow).DateTime:= QrAvance.FieldByName('dIdFecha').AsDateTime;
                Hoja.GetCellObject(2,IRow).SetCellText(QrAvance.FieldByName('sHoraInicio').AsString);
                Hoja.GetCellObject(3,IRow).SetCellText(QrAvance.FieldByName('sHoraFinal').AsString);
                Hoja.GetCellObject(4,IRow).SetCellText(QrAvance.FieldByName('sIdTipoMovimiento').AsString);
                Hoja.GetCellObject(5,IRow).SetCellText(sfnRestaHoras(QrAvance.FieldValues['sHoraFinal'], QrAvance.FieldValues['sHoraInicio']));
                Hoja.GetCellObject(6,IRow).SetCellText(FormatFloat( '0.00',QrAvance.FieldByName('AvanceAnterior').asfloat)  + '%');
                Hoja.GetCellObject(7,IRow).SetCellText(FormatFloat( '0.00',QrAvance.FieldByName('dAvance').AsFloat) + '%');
                Hoja.GetCellObject(8,IRow).SetCellText(FormatFloat( '0.00',QrAvance.FieldByName('dAvance').AsFloat + QrAvance.FieldByName('AvanceAnterior').AsFloat) + '%');

                TotalHrs:=sfnSumaHoras(TotalHrs,sfnRestaHoras(QrAvance.FieldByName('sHoraFinal').AsString, QrAvance.FieldByName('sHoraInicio').AsString));
                QrAvance.Next;
                BordeCelda(Hoja,1,8,IRow,0);
                inc(IRow);
                Hoja.Rows.ResetDefault(IRow);
              end;

              if QrAvance.RecordCount>0 then
              begin
               // inc(row);
                LockUnLock(Hoja,5,5,IRow,0,false);
                FormatoCelda(Hoja.GetCellObject(5,IRow),FrTitle);
                Hoja.GetCellObject(5,IRow).SetCellText(TotalHrs);
                AlineacionCelda(Hoja,haCENTER,vaCENTER,1,8,4,IRow-4);
                LockUnLock(Hoja,5,5,IRow,0);
               // Inc(IRow,4)
              end;
            end;
          end;
         // else
          //  Inc(IRow,5);

          IF (Hoja.Caption='COSTO POR ACTIVIDAD') then
          begin
            inc(IRow,3);
            Hoja.Rows.ResetDefault(IRow);
            ACellObj := GetCellObject(9,IRow);
            FormatoCelda(ACellObj,FrTitle);
            ACellObj.SetCellText('IMP MN');

            ACellObj := GetCellObject(10,IRow);
            FormatoCelda(ACellObj,Frtitle);
            ACellObj.SetCellText('IMP USD');

            AlineacionCelda(Hoja,haCENTER,vaCENTER,9,10,IRow,0);
            BordeCelda(Hoja,9,10,IRow,0);
            //Inc(IRow);
            pIRow:=IRow+2;
          end;

          IF (Hoja.Caption='ACTA DE ACTIVIDADES') then
          begin
            inc(IRow,1);
            Hoja.Rows.ResetDefault(IRow);
            ACellObj := GetCellObject(0,IRow);
            FormatoCelda(ACellObj,FrTitle);
            ACellObj.SetCellText('RESUMEN DE COSTOS:');
            RangeMerge(Hoja,0,9,IRow,0);
            Rows.Size[IRow] :=40;

            inc(IRow,1);
            Hoja.Rows.ResetDefault(IRow);
            ACellObj := GetCellObject(0,IRow);
            FormatoCelda(ACellObj,FrTitle);
            ACellObj.SetCellText('ANEXOS "C"');
            RangeMerge(Hoja,0,9,IRow,0);
            Rows.Size[IRow] :=30;

            inc(IRow,1);
            Hoja.Rows.ResetDefault(IRow);
            ACellObj := GetCellObject(0,IRow);
            FormatoCelda(ACellObj,FrTitle);
            ACellObj.SetCellText('ANEXO C');
            RangeMerge(Hoja,0,2,IRow,0);

            ACellObj := GetCellObject(3,IRow);
            FormatoCelda(ACellObj,FrTitle);
            ACellObj.SetCellText('DESCRIPCIÓN');
            RangeMerge(Hoja,3,7,IRow,0);

            ACellObj := GetCellObject(8,IRow);
            FormatoCelda(ACellObj,FrTitle);
            ACellObj.SetCellText('IMPORTE TOTAL');
            RangeMerge(Hoja,8,9,IRow,0);

            Rows.Size[IRow] :=25;

            inc(IRow,1);
            Hoja.Rows.ResetDefault(IRow);
            {ACellObj := GetCellObject(0,IRow);
            FormatoCelda(ACellObj,FrTitle);
            ACellObj.SetCellText('ANEXO C');   }
            RangeMerge(Hoja,0,2,IRow,0);

            ACellObj := GetCellObject(3,IRow);
            {FormatoCelda(ACellObj,FrTitle);
            ACellObj.SetCellText('DESCRIPCIÓN'); }
            RangeMerge(Hoja,3,7,IRow,0);

            ACellObj := GetCellObject(8,IRow);
            FormatoCelda(ACellObj,FrTitle);
            ACellObj.SetCellText('M.N.');

            ACellObj := GetCellObject(9,IRow);
            FormatoCelda(ACellObj,FrTitle);
            ACellObj.SetCellText('U.S.D.');


           // AlineacionCelda(Hoja,haCENTER,vaCENTER,9,10,IRow,0);
            //BordeCelda(Hoja,9,10,IRow,0);
            //Inc(IRow);
            pIRow:=IRow+1;
          end;





          while not QImportes.Eof do
          begin

              if QImportes.FieldByName('xRow').AsInteger=-1 then
                QImportes.Edit;

              IF (Hoja.Caption='COSTO POR ACTIVIDAD') or
                  (Hoja.Caption='ACTA DE ACTIVIDADES') then
                inc(iRow)
              else
                inc(iRow,2);

              Hoja.Rows.ResetDefault(IRow);
              if QImportes.fieldByName('eTipo').AsString<>'ACTIVIDAD' then
              begin
                if (Hoja.Caption='CAMPO') or (Hoja.Caption='DESGLOSE DE COSTOS') then
                begin
                  QRecursos.Active:=False;
                  QRecursos.ParamByName('Acta').AsInteger:=Datos.FieldByName('iIdActa').AsInteger;
                  QRecursos.ParamByName('wbs').AsString:=QrActividades.FieldByName('swbs').AsString;
                  QRecursos.ParamByName('Actividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
                  QRecursos.ParamByName('Tipo').AsString:=QImportes.FieldByName('eTipo').AsString;
                  QRecursos.Open;


                  if QImportes.FieldByName('eTipo').AsString='BARCO' then
                  begin
                    ACellObj := GetCellObject(0,IRow);
                    FormatoCelda(ACellObj,FrTitle);
                    ACellObj.SetCellText('MOVIMIENTO DE EMBARCACION');

                  end
                  else
                  begin
                    ACellObj := GetCellObject(0,IRow);
                    FormatoCelda(ACellObj,FrTitle);
                    ACellObj.SetCellText(QImportes.FieldByName('eTipo').AsString);
                  end;

                  LockUnLock(Hoja,0,10,IRow,0,False);
                  RangeMerge(Hoja,0,10,IRow,0);
                  LockUnLock(Hoja,0,10,IRow,0);
                  pIRow:=IRow;
                  inc(IRow);
                  Hoja.Rows.ResetDefault(IRow);
                  ACellObj := GetCellObject(0,IRow);
                  FormatoCelda(ACellObj,FrSubTitle);
                  ACellObj.SetCellText('PARTIDA');
                  ACellObj := GetCellObject(1,IRow);
                  FormatoCelda(ACellObj,FrSubTitle);
                  ACellObj.SetCellText('DESCRIPCIÓN');

                  if QImportes.FieldByName('eTipo').AsString='BARCO' then
                  begin
                    ACellObj := GetCellObject(5,IRow);
                    FormatoCelda(ACellObj,FrSubTitle);
                    ACellObj.SetCellText('CLAS.');
                  end
                  else
                  begin
                    ACellObj := GetCellObject(5,IRow);
                    FormatoCelda(ACellObj,FrSubTitle);
                    ACellObj.SetCellText('UNIDAD');
                  end;

                  ACellObj := GetCellObject(6,IRow);
                  FormatoCelda(ACellObj,FrSubTitle);
                  ACellObj.SetCellText('CANTIDAD');
                  ACellObj := GetCellObject(7,IRow);
                  FormatoCelda(ACellObj,FrSubTitle);
                  ACellObj.SetCellText('PU MN');
                  ACellObj := GetCellObject(8,IRow);
                  FormatoCelda(ACellObj,FrSubTitle);
                  ACellObj.SetCellText('PU USD');
                  ACellObj := GetCellObject(9,IRow);
                  FormatoCelda(ACellObj,FrSubTitle);
                  ACellObj.SetCellText('IMP MN');
                  ACellObj := GetCellObject(10,IRow);
                  FormatoCelda(ACellObj,FrSubTitle);
                  ACellObj.SetCellText('IMP USD');

                  LockUnLock(Hoja,1,4,IRow,0,False);
                  RangeMerge(Hoja,1,4,IRow,0);

                  while not QRecursos.Eof do
                  begin
                    Inc(IRow);
                    Hoja.Rows.ResetDefault(IRow);
                    if QRecursos.FieldByName('xRow').AsInteger=-1 then
                      QRecursos.Edit;
                    Hoja.GetCellObject(0,IRow).SetCellText(QRecursos.FieldByName('sIdRecurso').AsString);
                    AlineacionCelda(Hoja,haCENTER,vaCENTER,0,0,IRow,0);
                    Hoja.GetCellObject(1,IRow).SetCellText(QRecursos.FieldByName('mDescripcion').AsString);
                    Hoja.GetCellObject(5,IRow).SetCellText(QRecursos.FieldByName('sMedida').AsString);
                    AlineacionCelda(Hoja,haCENTER,vaCENTER,5,5,IRow,0);
                    Hoja.GetCellObject(6,IRow).SetCellText(FormatFloat( '0.00',QRecursos.FieldByName('dCantidad').Asfloat));
                    LockUnLock(Hoja,6,6,IRow,0,False);
                    Hoja.GetCellObject(7,IRow).SetCellText(FormatFloat( '0.00',QRecursos.FieldByName('dCostoMn').AsFloat));

                    Hoja.GetCellObject(8,IRow).SetCellText(FormatFloat( '0.00',QRecursos.FieldByName('dCostoDll').AsFloat));
                    if QRecursos.FieldByName('sFormulaMN').AsString='##' then
                    begin
                      Hoja.GetCellObject(9,IRow).Text:= '=Round((G' + IntToStr(IRow+1)+ ' * ' + 'H' + IntToStr(IRow+1)+ '),2)';
                      if QRecursos.State=dsEdit then
                        QRecursos.FieldByName('sFormulaMN').AsString:='=Round((G' + IntToStr(IRow+1)+ ' * ' + 'H' + IntToStr(IRow+1)+ '),2)';
                    end
                    else
                      Hoja.GetCellObject(9,IRow).Text:=QRecursos.FieldByName('sFormulaMN').AsString;

                    if QRecursos.FieldByName('sFormulaDll').AsString='##' then
                    begin
                      Hoja.GetCellObject(10,IRow).Text:= '=Round((G' + IntToStr(IRow+1)+ ' * ' + 'I' + IntToStr(IRow+1)+ '),2)' ;
                      if QRecursos.State=dsEdit then
                        QRecursos.FieldByName('sFormulaDll').AsString:='=Round((G' + IntToStr(IRow+1)+ ' * ' + 'I' + IntToStr(IRow+1)+ '),2)' ;
                    end
                    else
                      Hoja.GetCellObject(10,IRow).Text:=QRecursos.FieldByName('sFormulaDll').AsString;
                    RangeMerge(Hoja,1,4,IRow,0);

                    LockUnLock(Hoja,9,10,IRow,0,False);

                    if Length(QRecursos.FieldByName('mDescripcion').AsString)>200 then
                      Rows.Size[IRow] := Trunc((Length(QRecursos.FieldByName('mDescripcion').AsString)/60) * 20)
                    else
                      if Length(QRecursos.FieldByName('mDescripcion').AsString)>55 then
                        Rows.Size[IRow] := Trunc((Length(QRecursos.FieldByName('mDescripcion').AsString)/60) * 30);

                    if QRecursos.State=dsEdit then
                    begin
                      QRecursos.FieldByName('xRow').AsInteger:=IRow+1;
                      QRecursos.Post;
                    end;
                    QRecursos.Next;
                  end;


                  inc(iRow);
                  Hoja.Rows.ResetDefault(IRow);
                //IRow:=12;
                  ACellObj := GetCellObject(6, IRow);
                  FormatoCelda(ACellObj,FrTitle);
                  ACellObj.SetCellText(QImportes.FieldByName('mDescripcion').AsString);
                  LockUnLock(Hoja,6,8,IRow,0,False);
                  RangeMerge(Hoja,6,8,IRow,0);
                  LockUnLock(Hoja,6,8,IRow,0);
                  ACellObj := GetCellObject(9, IRow);
                  FormatoCelda(ACellObj,FrTitle);
                  if QImportes.FieldByName('sFormulaMn').AsString='##' then
                  begin
                    ACellObj.Text:= '=Round(sum(J' + IntToStr(pIRow+3)+ ':' + 'J' + IntToStr(IRow)+ '),2)' ;
                    if QImportes.State=dsEdit then
                      QImportes.FieldByName('sFormulaMn').AsString:='=Round(sum(J' + IntToStr(pIRow+3)+ ':' + 'J' + IntToStr(IRow)+ '),2)' ;
                  end
                  else
                    ACellObj.Text:=QImportes.FieldByName('sFormulaMn').AsString;

                  ACellObj := GetCellObject(10, IRow);
                  FormatoCelda(ACellObj,FrTitle);

                  if QImportes.FieldByName('sFormulaDll').AsString='##' then
                  begin
                    ACellObj.Text:= '=Round(sum(K' + IntToStr(pIRow+3)+ ':' + 'K' + IntToStr(IRow)+ '),2)' ;
                    if QImportes.State=dsEdit then
                      QImportes.FieldByName('sFormulaDll').AsString:='=Round(sum(K' + IntToStr(pIRow+3)+ ':' + 'K' + IntToStr(IRow)+ '),2)' ;
                  end
                  else
                    ACellObj.Text:=QImportes.FieldByName('sFormulaDll').AsString;

                  if QImportes.State=dsEdit then
                    QImportes.FieldByName('xRow').AsInteger:=IRow+1;
                  BordeCelda(Hoja,0,10,pIRow,IRow-pIRow-1);
                  BordeCelda(Hoja,6,10,IRow,0);
                  if sFormulaSumMn='' then
                     sFormulaSumMn:= 'J' + IntToStr(IRow+1)
                  else
                    sFormulaSumMn:= sFormulaSumMn +  '+J' + IntToStr(IRow+1);

                  if sFormulaSumDll='' then
                    sFormulaSumDll:='K' + IntToStr(IRow+1)
                  else
                    sFormulaSumDll:=sFormulaSumDll + '+K' + IntToStr(IRow+1);
                  AlineacionCelda(Hoja,haCENTER,vaCENTER,6,10,IRow,0);
                  LockUnLock(Hoja,9,10,IRow,0,False);
                  if  (QImportes.FieldByName('sLeyendaAnexo').AsString='') and
                      (QImportes.State = dsEdit) then
                    if Datos.FieldByName('eLugarOT').AsString='Tierra' then
                    begin
                      if QImportes.FieldByName('eTipo').AsString='BARCO' then
                        QImportes.FieldByName('sLeyendaAnexo').AsString:='ANEXO C 1.1';

                      if QImportes.FieldByName('eTipo').AsString='PERSONAL' then
                        QImportes.FieldByName('sLeyendaAnexo').AsString:='ANEXO C-5 "PERSONAL OPT."';

                      if QImportes.FieldByName('eTipo').AsString='EQUIPO' then
                        QImportes.FieldByName('sLeyendaAnexo').AsString:='ANEXO C-5 "EQUIPO OPT."';

                      if QImportes.FieldByName('eTipo').AsString='PERNOCTA' then
                        QImportes.FieldByName('sLeyendaAnexo').AsString:='ANEXO C-4 "SERV. DE HOTEL."';
                    end
                    else
                    begin
                      if QImportes.FieldByName('eTipo').AsString='BARCO' then
                        QImportes.FieldByName('sLeyendaAnexo').AsString:='ANEXO C 1.1';

                      if QImportes.FieldByName('eTipo').AsString='PERSONAL' then
                        QImportes.FieldByName('sLeyendaAnexo').AsString:='ANEXO C-2 "PERSONAL OPT."';

                      if QImportes.FieldByName('eTipo').AsString='EQUIPO' then
                        QImportes.FieldByName('sLeyendaAnexo').AsString:='ANEXO C-3 "EQUIPO OPT."';

                      if QImportes.FieldByName('eTipo').AsString='PERNOCTA' then
                        QImportes.FieldByName('sLeyendaAnexo').AsString:='ANEXO C-4 "SERV. DE HOTEL."';
                    end;

                end;

                if (Hoja.Caption='COSTO POR ACTIVIDAD')  then
                begin
                  ACellObj := GetCellObject(0, IRow);
                  FormatoCelda(ACellObj,FrContent);
                  if QImportes.FieldByName('eTipo').AsString='BARCO' then
                    ACellObj.SetCellText('MOVIMIENTO DE EMBARCACIÓN')
                  else
                    ACellObj.SetCellText(QImportes.FieldByName('eTipo').AsString);
                  RangeMerge(Hoja,0,5,IRow,0);

                  ACellObj := GetCellObject(6, IRow);
                  FormatoCelda(ACellObj,FrContent);
                  ACellObj.SetCellText(QImportes.FieldByName('mDescripcion').AsString);
                  LockUnLock(Hoja,6,8,IRow,0,False);
                  RangeMerge(Hoja,6,8,IRow,0);
                  LockUnLock(Hoja,6,8,IRow,0);
                  //AlineacionCelda(Hoja,haLeft,vaCENTER,5,5,IRow,0);


                  ACellObj := GetCellObject(9, IRow);
                  FormatoCelda(ACellObj,FrContent);
                  ACellObj.Text:='=CAMPO!J' + QImportes.FieldByName('xRow').AsString ;

                  ACellObj := GetCellObject(10, IRow);
                  FormatoCelda(ACellObj,FrContent);
                  ACellObj.Text:='=CAMPO!K' + QImportes.FieldByName('xRow').AsString ;

                  AlineacionCelda(Hoja,haLeft,vaCENTER,0,5,IRow,0);
                  AlineacionCelda(Hoja,haCENTER,vaCENTER,6,8,IRow,0);
                  BordeCelda(Hoja,0,10,IRow,0);
                  //Rows.Size[IRow] :=30;

                end;

                if (Hoja.Caption='ACTA DE ACTIVIDADES')  then
                begin
                  ACellObj := GetCellObject(0, IRow);
                  FormatoCelda(ACellObj,FrContent);
                  ACellObj.SetCellText(QImportes.FieldByName('sLeyendaAnexo').AsString);
                  RangeMerge(Hoja,0,2,IRow,0);

                  ACellObj := GetCellObject(3,IRow);
                  FormatoCelda(ACellObj,FrContent);
                  if QImportes.FieldByName('eTipo').AsString='BARCO' then
                    ACellObj.SetCellText('UTILIZANDO POSICIONAMIENTO DINÁMICO')
                  else
                    if QImportes.FieldByName('eTipo').AsString='PERNOCTA' then
                      ACellObj.SetCellText('SERVICIOS DE HOTELERIA')
                    else
                      ACellObj.SetCellText(QImportes.FieldByName('eTipo').AsString);
                  RangeMerge(Hoja,3,7,IRow,0);


                  ACellObj := GetCellObject(8, IRow);
                  FormatoCelda(ACellObj,FrContent);
                  ACellObj.Text:='='+quotedstr('DESGLOSE DE COSTOS')+'!J' + QImportes.FieldByName('xRow').AsString ;

                  ACellObj := GetCellObject(9, IRow);
                  FormatoCelda(ACellObj,FrContent);
                  ACellObj.Text:='='+quotedstr('DESGLOSE DE COSTOS')+'!K' + QImportes.FieldByName('xRow').AsString ;

                  AlineacionCelda(Hoja,haLeft,vaCENTER,0,7,IRow,0);
                  AlineacionCelda(Hoja,haRIGHT,vaCENTER,8,9,IRow,0);
                  //BordeCelda(Hoja,0,10,IRow,0);
                  //Rows.Size[IRow] :=30;

                end;



              end
              else
              begin
                if (Hoja.Caption='CAMPO') or (Hoja.Caption='DESGLOSE DE COSTOS') then
                begin
                  ACellObj := GetCellObject(5, IRow);
                  ACellObj.SetCellText('COSTO TOTAL DE LA ACTIVIDAD:');
                  LockUnLock(Hoja,5,8,IRow,0,False);
                  RangeMerge(Hoja,5,8,IRow,0);
                  LockUnLock(Hoja,5,8,IRow,0);
                  AlineacionCelda(Hoja,haLeft,vaCENTER,5,5,IRow,0);
                  ACellObj := GetCellObject(9, IRow);
                  FormatoCelda(ACellObj,FrTitle);
                  if QImportes.FieldByName('sFormulaMN').AsString='##' then
                  begin
                    ACellObj.Text:= '=' + sFormulaSumMn;
                    if QImportes.State=dsEdit then
                      QImportes.FieldByName('sFormulaMN').AsString:= '=' + sFormulaSumMn;
                  end
                  else
                    ACellObj.Text:=QImportes.FieldByName('sFormulaMN').AsString;

                  ACellObj := GetCellObject(10, IRow);
                  FormatoCelda(ACellObj,FrTitle);
                  if QImportes.FieldByName('sFormulaDll').AsString='##' then
                  begin
                    ACellObj.Text:= '='+  sFormulaSumDll;
                    if QImportes.State=dsEdit then
                      QImportes.FieldByName('sFormulaDll').AsString:= '='+  sFormulaSumDll;
                  end
                  else
                    ACellObj.Text:=QImportes.FieldByName('sFormulaDll').AsString;

                  AlineacionCelda(Hoja,haCENTER,vaCENTER,9,10,IRow,0);
                  BordeCelda(Hoja,5,10,IRow,0);
                  LockUnLock(Hoja,9,10,IRow,0,False);
                  Rows.Size[IRow] :=30;
                  if QImportes.State=dsEdit then
                    QImportes.FieldByName('xRow').AsInteger:=IRow+1;
                end;
                
                if (Hoja.Caption='COSTO POR ACTIVIDAD') then
                begin
                  inc(iRow);
                  Hoja.Rows.ResetDefault(IRow);
                  ACellObj := GetCellObject(9,IRow);
                  FormatoCelda(ACellObj,FrTitle);
                  ACellObj.SetCellText('IMP MN');

                  ACellObj := GetCellObject(10,IRow);
                  FormatoCelda(ACellObj,FrTitle);
                  ACellObj.SetCellText('IMP USD');

                  AlineacionCelda(Hoja,haCENTER,vaCENTER,9,10,IRow,0);
                  BordeCelda(Hoja,9,10,IRow,0);

                  Inc(iRow);
                  Hoja.Rows.ResetDefault(IRow);
                  ACellObj := GetCellObject(5, IRow);
                  FormatoCelda(ACellObj,FrTitle);
                  ACellObj.SetCellText('COSTO TOTAL DE LA ACTIVIDAD:');
                  LockUnLock(Hoja,5,8,IRow,0,False);
                  RangeMerge(Hoja,5,8,IRow,0);
                  LockUnLock(Hoja,5,8,IRow,0);
                  AlineacionCelda(Hoja,haRight,vaCENTER,5,10,IRow,0);
                  ACellObj := GetCellObject(9, IRow);
                  FormatoCelda(ACellObj,FrTitle);
                  ACellObj.Text:='=sum(J' + IntToStr(pIRow) + ':J' + IntToStr(IRow-2)+ ')';

                  ACellObj := GetCellObject(10, IRow);
                  FormatoCelda(ACellObj,FrTitle);
                  ACellObj.Text:='=sum(K' + IntToStr(pIRow) + ':K' + IntToStr(IRow-2)+ ')';
                  AlineacionCelda(Hoja,haCENTER,vaCENTER,9,10,IRow,0);
                  BordeCelda(Hoja,5,10,IRow,0);
                  Rows.Size[IRow] :=30;
                end;

                if (Hoja.Caption='ACTA DE ACTIVIDADES')  then
                begin
                  RangeMerge(Hoja,0,2,IRow,0);

                  ACellObj := GetCellObject(3,IRow);
                  FormatoCelda(ACellObj,FrSubTitle);
                  ACellObj.SetCellText('TOTALES');
                  RangeMerge(Hoja,3,7,IRow,0);

                  ACellObj := GetCellObject(8,IRow);
                  FormatoCelda(ACellObj,FrTitle);
                  ACellObj.Text:='=sum(I' + IntToStr(pIRow+1) + ':I' + IntToStr(IRow)+ ')';

                  ACellObj := GetCellObject(9,IRow);
                  FormatoCelda(ACellObj,FrTitle);
                  ACellObj.Text:='=sum(J' + IntToStr(pIRow+1) + ':J' + IntToStr(IRow)+ ')';

                  BordeCelda(Hoja,0,9,pIRow-4,IRow-(pIRow-4));
                end;
              end;
            //IRow:=NPdas;
            {LockUnLock(Hoja,0,1,IRow,0,False);
            RangeMerge(Hoja,0,1,IRow,0);
            LockUnLock(Hoja,0,1,IRow,0);

            LockUnLock(Hoja,2,10,IRow,0,False);
            RangeMerge(Hoja,2,10,IRow,0);
            LockUnLock(Hoja,2,10,IRow,0);

            Inc(IRow);
            LockUnLock(Hoja,0,1,IRow,0,False);
            RangeMerge(Hoja,0,1,IRow,0);
            LockUnLock(Hoja,0,1,IRow,0);

            LockUnLock(Hoja,2,10,IRow,0,False);
            RangeMerge(Hoja,2,10,IRow,0);
            LockUnLock(Hoja,2,10,IRow,0);

            Inc(IRow,2);
            LockUnLock(Hoja,1,8,IRow,0,False);
            RangeMerge(Hoja,1,8,IRow,0);
            LockUnLock(Hoja,1,8,IRow,0);


            Inc(IRow,3);
            LockUnLock(Hoja,1,4,IRow,0,False);
            RangeMerge(Hoja,1,4,IRow,0);
            LockUnLock(Hoja,1,4,IRow,0);

            Inc(IRow,2);
            LockUnLock(Hoja,0,10,IRow,0,False);
            RangeMerge(Hoja,0,10,IRow,0);
            LockUnLock(Hoja,0,10,IRow,0);

            Inc(IRow);
            LockUnLock(Hoja,1,4,IRow,0,False);
            RangeMerge(Hoja,1,4,IRow,0);

            Inc(IRow);
            LockUnLock(Hoja,1,4,IRow,0,False);
            RangeMerge(Hoja,1,4,IRow,0);

            Inc(IRow);
            LockUnLock(Hoja,6,8,IRow,0,False);
            RangeMerge(Hoja,6,8,IRow,0);
            LockUnLock(Hoja,6,8,IRow,0);

            IRow:=NPdas;
            BordeCelda(Hoja,0,10,IRow,1);

           // IRow:=3;
            Inc(IRow,3);
            BordeCelda(Hoja,1,8,IRow,2);



            //IRow:=5;
            Inc(IRow,3);
            BordeCelda(Hoja,1,5,IRow,0);  

            //IRow:=7;
            Inc(IRow,2);
            BordeCelda(Hoja,0,10,IRow,2);

           // IRow:=10;
            Inc(IRow,3);
            BordeCelda(Hoja,6,10,IRow,0);

            IIRow:=12;
            BordeCelda(Hoja,0,10,IRow,2);
            IRow:=15;
            BordeCelda(Hoja,6,10,IRow,0);

            nc(IRow,3);
            LockUnLock(Hoja,1,4,IRow,0,False);
            RangeMerge(Hoja,1,4,IRow,0);
            LockUnLock(Hoja,1,4,IRow,0);






            LockUnLock(Hoja,0,10,12,0,False);
            RangeMerge(Hoja,0,10,12,0);
            LockUnLock(Hoja,0,10,12,0);

            LockUnLock(Hoja,1,4,13,0,False);
            RangeMerge(Hoja,1,4,13,0);
            LockUnLock(Hoja,1,4,13,0);

            LockUnLock(Hoja,1,4,14,0,False);
            RangeMerge(Hoja,1,4,14,0);

            LockUnLock(Hoja,6,8,15,0,False);
            RangeMerge(Hoja,6,8,15,0);
            LockUnLock(Hoja,6,8,15,0);

            }

            if QImportes.State=dsEdit then
              QImportes.Post;
            QImportes.Next;
          end;
        end;
        QrActividades.Next;
      end;




    end;

  finally
    QrActividades.Destroy;
    QImportes.Destroy;
    QRecursos.Destroy;
    QrAvance.Destroy;
  end;
end;
*)
procedure TFrmConsultas.ModificarEstructura(FechaI,FechaT:TDate;Tipo: FtConsulta);

begin


end;

procedure TFrmConsultas.ConsultaAvances(FechaI: TDate; FechaT: TDate;Tipo:FtConsulta=FtDiario;Folios: string = '');
var
  QrFolios:TZReadOnlyQuery;
 // QrAvance
begin
  QrFolios:=TZReadOnlyQuery.Create(nil);
  try
    QrFolios.SQL.Text:= 'select ot.sNumeroOrden,ot.sidFolio,ot.mdescripcion,' + #13#10 +
                        'xRound(AvancesAnteriores(:FechaI,ot.sContrato,ot.sNumeroOrden),2) as AvanceAnterior from ordenesdetrabajo ot' + #13#10 + 
                        'inner join bitacoradeactividades ba' + #13#10 + 
                        'on(ba.sContrato=ot.sContrato and ba.sNumeroOrden=ot.sNumeroOrden)' + #13#10 + 
                        'where ot.sContrato=:Contrato and (:Folio=-1 or (:Folio<>-1 and FIND_IN_SET(ot.sIdFolio,:Folio)))' + #13#10 + 
                        'and ba.didFecha between :FechaI and :FechaT' + #13#10 +
                        'group by ot.sNumeroOrden order by ot.iOrden';
    QrFolios.ParamByName('Contrato').AsString:=global_contrato;
    if Folios='' then
      QrFolios.ParamByName('Folio').AsInteger:=-1
    else
      QrFolios.ParamByName('Folio').AsString:=Folios;
    QrFolios.ParamByName('FechaI').AsDate:=FechaI;
    QrFolios.ParamByName('FechaT').AsDate:=FechaT;
    QrFolios.Open;
    while not QrFolios.Eof do
    begin


      QrFolios.Next;
    end;
  finally
    QrFolios.Destroy;
  end;

end;







end.
