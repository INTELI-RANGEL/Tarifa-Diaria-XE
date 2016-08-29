unit UTSuperPanel;

interface
uses SysUtils,Graphics,Controls,StdCtrls,ExtCtrls,Forms,NxGrid,NxCustomGridControl,
      propscrl,Types,AdvGrid,NxColumnClasses,Classes,NxColumns,Grids;

type
  TSuperPanel=class
    private
      FsContrato,FsNumeroOrden:string;
      FIdentificador:string;
      iAuxPos,RowAnt,
      ColAnt,LastPos:Integer;
      isRecursive:Boolean;
      FListaPdas:TStringList;
      FListaRecurso:TStringList;
      procedure MiScroll(Sender: TObject;ScrollCode: TScrollCode; var ScrollPos: Integer);
      procedure MiEditAccept(Sender: TObject; ACol,ARow: Integer; Value: WideString; var Accept: Boolean);
      procedure MiSelectCell(Sender: TObject; ACol,ARow: Integer);
      procedure DatosVerticalScroll(Sender: TObject;Position: Integer);
      procedure TotalesVerticalScroll(Sender: TObject;Position: Integer);
      //procedure MiSelectCell(Sender: TObject; ACol, ARow: Integer);
     
    public
      PnlMain:Tpanel;
      PnlDatos:TPanel;
      PnlTotales:TPanel;
      PnlEncDatos:TPanel;
      PnlEncTotales:TPanel;
      PnlScroll:TPanel;
      PnlSubScroll:TPanel;
      PnlRightScroll:TPanel;
      NxGridDatos:TNextGrid;
      NxGridTotales:TNextGrid;
      PSbDatos:TPropScrollbar;
      AvGridTurnos:TAdvStringGrid;
      AvGridPdas:TAdvStringGrid;
      AvGrid1:TAdvStringGrid;
      AvGrid2:TAdvStringGrid;
      Constructor Create(Sender: TWinControl; sParamContrato,Identificador: String);
      destructor Destroy;
    published
      property ListaPdas:TStringList read FListaPdas write FListaPdas;
      property ListaRecursos:TStringList read FListaRecurso write FListaRecurso;
      property Id:string Read FIdentificador;
      property sContrato:string read FsContrato;
      property sNumeroOrden:string read FsNumeroOrden;
  end;

  TPartida=class
    protected
      FidDiario:Integer;
      FiCol:Integer;
      FsPartida:string;
      FsIdClasificacion:string;
      FsHoraInicio:string;
      FsHoraFinal:string;

    public
      property IdDiario:Integer read FidDiario write FidDiario;
      property iCol:Integer read FiCol write FiCol;
      Property sPartida:string read FsPartida write FsPartida;
      Property sIdClasificacion:string read FsIdClasificacion write FsIdClasificacion;
      Property sHoraInicio:string read FsHoraInicio write FsHoraInicio;
      property sHoraFinal:string read FsHoraFinal write FsHoraFinal;
  end;

  TRecurso=class
    protected
      FIdRecurso:string;
      FDescripcion:string;
      FItemOrden:Integer;
      FsPlataforma:string;
      FsAgrupa:string;
      FiReng:Integer;
    public
      Property sIdRecurso:string read FIdRecurso write FIdRecurso;
      property sDescripcion:string read FDescripcion write FDescripcion;
      property ItemOrden:Integer read FItemOrden write FItemOrden;
      Property sPlataforma:string read FsPlataforma write FsPlataforma;
      property sAgrupa:string read FsAgrupa write FsAgrupa;
      property iReng:Integer read FiReng write FiReng;
  end;

implementation

constructor TSuperPanel.Create(Sender: TWinControl; sParamContrato,Identificador: string);
var
  ColT:TNxTextColumn;
  sNombre:string;
begin
  FsContrato:=sParamContrato;
  FsNumeroOrden:=Identificador;
  FIdentificador:=Identificador;
  FListaPdas:=TStringList.Create;
  FListaRecurso:=TStringList.Create;
  PnlMain:=TPanel.Create(Sender);
  with PnlMain do
  begin
    Parent:=Sender;
    Align:=alClient;
    Caption:='';
    BevelEdges:=[];
    BevelInner:=bvNone;
    BevelOuter:=bvNone;
    BorderStyle:=bsNone;
    BevelKind:=bkNone;
  end;

  PnlDatos:=TPanel.Create(sender);
  with PnlDatos do
  begin
    Parent:=PnlMain;
    Align:=alClient;
    Caption:='';
    BevelEdges:=[];
    BevelInner:=bvNone;
    BevelKind:=bkNone;
    BevelOuter:=bvNone;
    BorderStyle:=bsNone;
  end;

  PnlTotales:=TPanel.Create(sender);
  with PnlTotales do
  begin
    Parent:=PnlMain;
    Align:=alRight;
    Caption:='';
   { BevelEdges:=[];    }
    BevelInner:=bvNone;
    BevelKind:=bkNone;
    BevelOuter:=bvLowered;
    BorderStyle:=bsNone;
    Width:=82;
  end;

  PnlEncDatos:=TPanel.Create(sender);
  with PnlEncDatos do
  begin
    Parent:=PnlDatos;
    Align:=alTop;
    Caption:='';
    BevelEdges:=[];
    BevelInner:=bvNone;
    BevelKind:=bkNone;
    BevelOuter:=bvNone;
    BorderStyle:=bsNone;
    Height:=44;
  end;

  PnlEncTotales:=TPanel.Create(sender);
  with PnlEncTotales do
  begin
    Parent:=PnlTotales;
    Align:=alTop;
    Caption:='';
    BevelEdges:=[];
    BevelInner:=bvNone;
    BevelKind:=bkNone;
    BevelOuter:=bvNone;
    BorderStyle:=bsNone;
    Height:=43;
  end;

  PnlScroll:=TPanel.Create(sender);
  with PnlScroll do
  begin
    Parent:=PnlMain;
    Align:=alBottom;
    Caption:='';
    BevelEdges:=[];
    BevelInner:=bvNone;
    BevelKind:=bkNone;
    BevelOuter:=bvNone;
    BorderStyle:=bsNone;
    Height:=19;
  end;

  PnlSubScroll:=TPanel.Create(sender);
  with PnlSubScroll do
  begin
    Parent:=PnlScroll;
    Align:=alClient;
    Caption:='';
    BevelEdges:=[];
    BevelInner:=bvNone;
    BevelKind:=bkNone;
    BevelOuter:=bvNone;
    BorderStyle:=bsNone;
  end;

  PnlRightScroll:=TPanel.Create(sender);
  with PnlRightScroll do
  begin
    Parent:=PnlScroll;
    Align:=alRight;
    Caption:='';
    BevelEdges:=[];
    BevelInner:=bvNone;
    BevelKind:=bkNone;
    BevelOuter:=bvNone;
    BorderStyle:=bsNone;
    Width:=82;
  end;

  PSbDatos:=TPropScrollbar.Create(sender);
  with PSbDatos do
  begin
    Parent:=PnlSubScroll;
    Align:=alTop;
    OnScroll:=MiScroll;
  end;

  NxGridDatos:=TNextGrid.Create(Sender);
  with NxGridDatos do
  begin
    sNombre:='NxGrid'+StringReplace(FIdentificador,' ','',[rfReplaceAll]);
    sNombre:=StringReplace(sNombre,'-','_',[rfReplaceAll]);
    NxGridDatos.Name:=sNombre;
    Parent:=PnlDatos;
    Align:=alClient;
    AutoScroll:=False;
    AppearanceOptions:=[aoHighlightSlideCells,aoIndicateSelectedCell];
    HeaderStyle:=hsFlatBorders;
    HideScrollBar:=True;
    Options:=[goGrid,goHeader,goIndicator,goSecondClickEdit,goSelectFullRow];
    BorderStyle:=bsNone;
    HighlightedTextColor:=clred;
    InactiveSelectionColor:=clHighlight;
    SelectionColor:=clHighlight;
    OnEditAccept:=MiEditAccept;
    OnSelectCell:=MiSelectCell;
    OnVerticalScroll:=DatosVerticalScroll;
  end;

  NxGridTotales:=TNextGrid.Create(sender);
  with NxGridTotales do
  begin
    Parent:=PnlTotales;
    Align:=alClient;
    AppearanceOptions:=[aoHideSelection,aoHighlightSlideCells,aoIndicateSelectedCell];
    HeaderStyle:=hsFlatBorders;
    HideScrollBar:=True;
    Options:=[goDisableColumnMoving,goGrid,goHeader,goSecondClickEdit,goSelectFullRow];
    BorderStyle:=bsNone;
    HighlightedTextColor:=clred;
    InactiveSelectionColor:=clHighlight;
    SelectionColor:=clHighlight;
    ColT:=TNxTextColumn.Create(Sender);
    ColT.Header.Caption:='TOTALES';
    ColT.Header.Alignment:=Talignment(2);
    ColT.Options:=[coCanClick,coCanInput,coCanSort,coPublicUsing,coShowTextFitHint];
    ColT.Alignment:=Talignment(2);
    ColT.Width:=60;
    ColT.Color:=$00D8D8D8;
    ColT.Font.Style:=[fsBold];

    Columns.AddColumn(ColT);
    OnVerticalScroll:=TotalesVerticalScroll;
  end;



  AvGridPdas:=TAdvStringGrid.Create(Sender);
  with AvGridPdas do
  begin
    Parent:=PnlEncDatos;
    ColCount:=2;
    FixedCols:=0;
    FixedRows:=1;
    Flat:=True;
    Look:=glStandard;
    RowCount:=2;
    ScrollBarAlways:=saNone;
    ScrollBars:=ssnone;
    Align:=alTop;
    BevelEdges:=[];
    BevelInner:=bvNone;
    BevelKind:=bkNone;
    BevelOuter:=bvNone;
    BorderStyle:=bsNone;
    Color:=$00FFAE5E;
    FixedColor:=$00FFAE5E;

    Height:=22;
    DefaultAlignment:=taCenter;
  end;

  AvGridTurnos:=TAdvStringGrid.Create(Sender);
  with AvGridTurnos do
  begin
    Parent:=PnlEncDatos;
    ColCount:=2;
    FixedCols:=0;
    FixedRows:=1;
    Flat:=True;
    Look:=glStandard;
    RowCount:=2;
    ScrollBarAlways:=saNone;
    ScrollBars:=ssnone;
    Align:=alTop;
    BevelEdges:=[];
    BevelInner:=bvNone;
    BevelKind:=bkNone;
    BevelOuter:=bvNone;
    BorderStyle:=bsNone;
    Color:= $00ABF18D;
    FixedColor:=$00ABF18D;
    Height:=22;
    Font.Height:=-10;
    DefaultAlignment:=taCenter;
  end;



  AvGrid2:=TAdvStringGrid.Create(Sender);
  with AvGrid2 do
  begin
    Parent:=PnlEncTotales;
    ColCount:=2;
    FixedCols:=0;
    FixedRows:=1;
    Flat:=True;
    Look:=glStandard;
    RowCount:=2;
    ScrollBarAlways:=saNone;
    ScrollBars:=ssnone;
    Align:=alTop;
    BevelEdges:=[];
    BevelInner:=bvNone;
    BevelKind:=bkNone;
    BevelOuter:=bvNone;
    BorderStyle:=bsNone;
    Color:=$00FFAE5E;
    FixedColor:=$00FFAE5E;
    Height:=22;
    Options:=[goRangeSelect];
  end;

  AvGrid1:=TAdvStringGrid.Create(Sender);
  with AvGrid1 do
  begin
    Parent:=PnlEncTotales;
    ColCount:=2;
    FixedCols:=0;
    FixedRows:=1;
    Flat:=True;
    Look:=glStandard;
    RowCount:=2;
    ScrollBarAlways:=saNone;
    ScrollBars:=ssnone;
    Align:=alTop;
    BevelEdges:=[];
    BevelInner:=bvNone;
    BevelKind:=bkNone;
    BevelOuter:=bvNone;
    BorderStyle:=bsNone;
    Color:= $00ABF18D;
    FixedColor:=$00ABF18D;
    Height:=20;
    Options:=[goRangeSelect];
  end;
end;

destructor TSuperPanel.Destroy;
begin
  FreeAndNil(AvGrid1);
  FreeAndNil(AvGrid2);
  FreeAndNil(AvGridTurnos);
  FreeAndNil(AvGridPdas);
  FreeAndNil(PSbDatos);
  FreeAndNil(NxGridDatos);
  FreeAndNil(NxGridTotales);
  FreeAndNil(PnlRightScroll);
  FreeAndNil(PnlSubScroll);
  FreeAndNil(PnlScroll);
  FreeAndNil(PnlEncDatos);
  FreeAndNil(PnlEncTotales);
  FreeAndNil(PnlTotales);
  FreeAndNil(PnlDatos);
  FreeAndNil(PnlMain);

  inherited;
end;

procedure TSuperPanel.MiScroll(Sender: TObject;
  ScrollCode: TScrollCode; var ScrollPos: Integer);
var
  dif,AuxDif:Integer;
  iCicle:Integer;
  iAvance:Double;
begin
  dif:=NxGridDatos.Columns.Count - 13;

  if ScrollCode in [scLineDown,scPageDown] then
  begin

    iCicle:=0;
    while iCicle<14 do
    begin
      NxGridDatos.HorzScrollBar.Next;
      Inc(iCicle);
    end;
            //(ScrollPos-1) +

    if AvGridTurnos.Col<12 then
    begin
      AvGridTurnos.Col:=(AvGridTurnos.Col + 12)-1;
      AvGridPdas.Col:=(AvGridPdas.Col + 12)-1;

    end
    else
    begin
      if (AvGridTurnos.Col + 1)<AvGridTurnos.ColCount then
      begin
        AvGridTurnos.Col:=AvGridTurnos.Col + 1;
        AvGridPdas.Col:=AvGridPdas.Col + 1;
      end;
    
    end;
    //ScrollPos:= (Trunc((((AdGridTurnos.Col)-12)*100/dif)));
  end;

  if ScrollCode in [scLineUp,scPageUp] then
  begin
    iCicle:=0;
    while iCicle<14 do
    begin
      NxGridDatos.HorzScrollBar.Prior;
      Inc(iCicle);
    end;
              //(ScrollPos+1) -
    if (AvGridTurnos.Col-1)>1 then
    begin
      if AvGridTurnos.Col>12 then
      begin
        AvGridTurnos.Col:=(AvGridTurnos.Col - 12)+1;
        AvGridPdas.Col:=(AvGridPdas.Col - 12)+1;
        AuxDif:=(AvGridPdas.Col-2);
      end
      else
      begin
        AvGridTurnos.Col:=AvGridTurnos.Col - 1;
        AvGridPdas.Col:=AvGridPdas.Col - 1;
        AuxDif:=(AvGridPdas.Col-2);
      end;
    end;
    //ScrollPos:= (Trunc(((AuxDif)*100/dif)));
  end;

  if ScrollCode=scTrack then
  begin
    //iAvance:=(dif/100);
    //AdGridTurnos.
    if iAuxPos<ScrollPos then
    begin
      iCicle:=0;
      while iCicle<14 do
      begin
        NxGridDatos.HorzScrollBar.Next;
        Inc(iCicle);
      end;
              //(ScrollPos-1) +

      if AvGridTurnos.Col<12 then
      begin
        AvGridTurnos.Col:=(AvGridTurnos.Col + 12)-1;
        AvGridPdas.Col:=(AvGridPdas.Col + 12)-1;

      end
      else
      begin
        if (AvGridTurnos.Col + 1)<AvGridTurnos.ColCount then
        begin
          AvGridTurnos.Col:=AvGridTurnos.Col + 1;
          AvGridPdas.Col:=AvGridPdas.Col + 1;
        end;
      end;
      iAuxPos:=ScrollPos;
    end
    else
    begin
      if iAuxPos>ScrollPos then
      begin
        iCicle:=0;
        while iCicle<14 do
        begin
          NxGridDatos.HorzScrollBar.Prior;
          Inc(iCicle);
        end;
        
        if (AvGridTurnos.Col-1)>1 then
        begin
          if AvGridTurnos.Col>12 then
          begin
            AvGridTurnos.Col:=(AvGridTurnos.Col - 12)+1;
            AvGridPdas.Col:=(AvGridPdas.Col - 12)+1;
            AuxDif:=(AvGridPdas.Col-2);
          end
          else
          begin
            AvGridTurnos.Col:=AvGridTurnos.Col - 1;
            AvGridPdas.Col:=AvGridPdas.Col - 1;
            AuxDif:=(AvGridPdas.Col-2);
          end;
        end;
        iAuxPos:=ScrollPos;
      end;

    end;
  
  end;
end;

procedure TSuperPanel.MiSelectCell(Sender: TObject; ACol,
  ARow: Integer);
var
  dif,AuxDif:Integer;
begin

  if RowAnt<>ARow then
  begin
    NxGridTotales.Cell[0,RowAnt].Color:=$00D8D8D8;
    NxGridTotales.Cell[0,ARow].Color:=clHighlight;
    RowAnt:=ARow;
  end;

  if NxGridDatos.Columns.Count>13 then
  begin
    dif:=NxGridDatos.Columns.Count - 13;
    if (ACol>12) or (ACol<=(NxGridDatos.Columns.Count-12)) then
    begin
      //if AvGridTurnos.Col<12 then
      //begin
        AvGridTurnos.Col:=Acol;//(AvGridTurnos.Col + 12)-1;
        AvGridPdas.Col:=Acol;//(AvGridPdas.Col + 12)-1;

        if ColAnt<ACol then
        begin
          PSbDatos.Position:=(Acol-12);
          LastPos:=ACol;
        end
        else
        begin
         // if (ACol<=(NxGridDatos.Columns.Count-12)) and (ACol>1) then
          if (ACol<=(LastPos-11)) and (ACol>1) then
            PSbDatos.Position:=Acol - 2;
        end;

        
        ColAnt:=ACol;
        {if ACol>12 then
          PSbDatos.Position:=(Acol-12)
        else
          PSbDatos.Position:=(dif-Acol); }

          //PSbDatos.Position:=NxGridDatos.HorzScrollBar.Position;

                                                //end;
      {
      else
      begin
        if (AvGridTurnos.Col + 1)<AvGridTurnos.ColCount then
        begin
          AvGridTurnos.Col:=AvGridTurnos.Col + 1;
          AvGridPdas.Col:=AvGridPdas.Col + 1;
        end;
      end;}
    end;

  end;
 
    //AvGridTurnos.Col:=(AvGridTurnos.Col + 12)-1;
    //AvGridPdas.Col:=(AvGridPdas.Col + 12)-1;

end;

procedure TSuperPanel.MiEditAccept(Sender: TObject; ACol,
  ARow: Integer; Value: WideString; var Accept: Boolean);
var
  dTotalDia:Double;
  i:Integer;
  dDia:Integer;
begin
  if TryStrToInt(Value,i) then
  begin
    dDia:=StrToIntDef(NxGridDatos.Cell[Acol,Arow].AsString,0);
    ddia:=(StrToIntDef(Value,0)-ddia);
    dTotalDia:=0;
    for I := 2 to NxGridDatos.Columns.Count - 1 do
      dTotalDia:=dTotalDia +   StrToFloatDef(NxGridDatos.Cell[i,ARow].AsString,0);
    NxGridTotales.Cell[0,ARow].AsFloat:=dTotalDia+ddia;
  end
  else
    Accept:=false;
end;

procedure TSuperPanel.DatosVerticalScroll(Sender: TObject;
  Position: Integer);
begin
  if not isRecursive then
  begin
    isRecursive:=True;
    NxGridTotales.VertScrollBar.Position:=Position;
    isRecursive:=False;
  end;
end;

procedure TSuperPanel.TotalesVerticalScroll(Sender: TObject;
  Position: Integer);
begin
  if not isRecursive then
  begin
    isRecursive:=True;
    NxGridDatos.VertScrollBar.Position:=Position;
    isRecursive:=False;
  end;
end;



end.
