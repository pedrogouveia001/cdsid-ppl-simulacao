unit UntEntrada;

interface

uses
 Classes, SysUtils, IWAppForm, IWApplication, IWColor, IWTypes, IWCompButton,
  IWCompListbox, IWCompEdit, IWVCLBaseControl, IWBaseControl, IWBaseHTMLControl,
  IWControl, IWCompLabel, Controls, Forms, IWVCLBaseContainer, IWContainer,
  IWHTMLContainer, IWHTML40Container, IWRegion, IWGrids, ComObj, Dialogs,
  Variants, Math, UserSessionUnit, IWExtCtrls, jpeg, windows, ActiveX,
  IWWebGrid, IWAdvWebGrid, IWCompTabControl, IWAdvWebGridExcel, IWCompText,
  IWCompCheckbox, IWBaseComponent, IWBaseHTMLComponent, IWBaseHTML40Component,
  chart, series, pngimage, Grids,
  IWCompProgressBar, ExtCtrls, IWHTMLControls, TeeGDIPlus, TeEngine, TeeProcs,
  IWTMSPopup, IWHelpTip, Menus,

  BubbleCh, TeePolar,
  IWCompRadioButton, IWTMSCheckList, TeeLisB,
  TeeSeriesTextEd, Buttons, CheckLst,
  ShellApi, TeeLegendScrollBar, Spin, ToolWin, StdActns,
  TeeURL, TeeXML, AppEvnts, IWCompDynamicChart, IWVCLComponent, TeeProCo,
  Messages, Graphics, ComCtrls, StdCtrls, ErrorBar,
  IWProgressIndicator, IWDBGrids,
  IWColorComboBox, IWTMSImgCtrls, IWTMSCtrls,
  IWTMSMenus, RpDefine, RpRender, RpRenderText,
  IWAdvCheckGroup, IWDBAdvCheckGroup, GanttCh, IWCompRectangle,
  TeeDoubleHorizBar;
type
  TIWForm1 = class(TIWAppForm)
    rgnmatcons: TIWRegion;
    IWLabel49: TIWLabel;
    IWImage31: TIWImage;
    IWImage32: TIWImage;
    IWLabel20: TIWLabel;
    IWRegion1: TIWRegion;
    IWLabel3: TIWLabel;
    IWLabel1: TIWLabel;
    cmbbxcriteria: TIWComboBox;
    IWLabel2: TIWLabel;
    btnshowresults: TIWButton;
    rgnresultado: TIWRegion;
    IWImage1: TIWImage;
    IWImage2: TIWImage;
    IWLabel7: TIWLabel;
    chrtResultado: TChart;
    rgnprimeiro: TIWRegion;
    grdresultados: TIWGrid;
    lbltexto: TIWLabel;
    imgResultado: TIWImage;
    Chart1: TChart;
    Series2: TBarSeries;
    edtnumalt: TIWEdit;
    grdImport: TTIWAdvWebGrid;
    TIWAdvWebGridExcelIO1: TTIWAdvWebGridExcelIO;
    IWRegion2: TIWRegion;
    grdnsolutions: TIWGrid;
    IWLabel4: TIWLabel;
    IWLink1: TIWLink;
    IWLink2: TIWLink;
    IWLink3: TIWLink;
    rgnestatistica: TIWRegion;
    IWLabel5: TIWLabel;
    IWLabel6: TIWLabel;
    IWLabel9: TIWLabel;
    IWLabel10: TIWLabel;
    lblmedia: TIWLabel;
    lblmediana: TIWLabel;
    lblmoda: TIWLabel;
    lbldesviopadrao: TIWLabel;
    IWRegion5: TIWRegion;
    IWImage3: TIWImage;
    IWRegion7: TIWRegion;
    Series1: TBarSeries;
    IWImage4: TIWImage;
    IWRegion3: TIWRegion;
    IWRegion4: TIWRegion;
    rgnresultadoPromethee: TIWRegion;
    IWImage5: TIWImage;
    IWImage6: TIWImage;
    IWLabel11: TIWLabel;
    IWLabel12: TIWLabel;
    imgresultadopromethee: TIWImage;
    IWLabel13: TIWLabel;
    IWLink4: TIWLink;
    IWLink5: TIWLink;
    IWImage8: TIWImage;
    rgnestatisticapromethee: TIWRegion;
    IWLabel14: TIWLabel;
    IWLabel15: TIWLabel;
    IWLabel16: TIWLabel;
    IWLabel17: TIWLabel;
    lblmediapromethee: TIWLabel;
    lblmedianapromethee: TIWLabel;
    lblmodapromethee: TIWLabel;
    lbldesviopadraopromethee: TIWLabel;
    chrtresultadopromethee: TChart;
    BarSeries1: TBarSeries;
    IWRegion10: TIWRegion;
    grdnsolutionsPromethee: TIWGrid;
    IWRegion11: TIWRegion;
    IWRegion12: TIWRegion;
    rgnregradecisao: TIWRegion;
    grdregradecisao: TIWGrid;
    IWLink7: TIWLink;
    rgnregradecisaoPromethee: TIWRegion;
    grdregradecisaopromethee: TIWGrid;
    IWLink6: TIWLink;
    IWLabel41: TIWLabel;
    IWLabel18: TIWLabel;
    IWLabel19: TIWLabel;
    procedure cmbbxauxiliarAsyncChange(Sender: TObject;
      EventParams: TStringList);
    procedure cmbbxauxiliarChange(Sender: TObject);
    procedure btnshowresultsClick(Sender: TObject);
    procedure GerarCases;
    procedure rgnresultadoCreate(Sender: TObject);
    procedure SetarCores;
    procedure rgnmatconsCreate(Sender: TObject);
    function Mean(const Data: array of Double): Double;
    function StandardDeviation(const Data: array of Double; IsSample: Boolean)
      : Double;
    procedure IWLink1Click(Sender: TObject);
    procedure LimpaVariaveis;
    procedure IWLink2Click(Sender: TObject);
    procedure IWLink3Click(Sender: TObject);
    procedure rgnresultadoPrometheeCreate(Sender: TObject);
    procedure IWLink4Click(Sender: TObject);
    procedure IWLink7Click(Sender: TObject);
    procedure IWLink6Click(Sender: TObject);
    procedure PLEscolha(auxinst: Integer);
    procedure DSiTrimWorkingSet;
    function TrocaVP(Num: Double): String;
  public
  end;

implementation

uses lpsolve51, uAPI_SolveLP, ServerController;

{$R *.dfm}

Function TIWForm1.TrocaVP(Num: Double): String;
var
  S: string;
  i: Integer;
begin
  S := FloatToStr(Num);
  for i := 2 to length(S) - 1 do
    if S[i] = ',' then
      S[i] := '.';

  Result := S;
end;

function TIWForm1.Mean(const Data: array of Double): Double;
var
  Sum: Double;
  I: Integer;
begin
  Sum := 0;
  for I := Low(Data) to High(Data) do
    Sum := Sum + Data[I];
  Result := Sum / Length(Data);
end;

function TIWForm1.StandardDeviation(const Data: array of Double;
  IsSample: Boolean): Double;
var
  MeanValue, Sum: Double;
  I: Integer;
  N: Integer;
begin
  if Length(Data) = 0 then
    raise Exception.Create('Data array is empty.');

  MeanValue := Mean(Data);
  Sum := 0;

  for I := Low(Data) to High(Data) do
    Sum := Sum + Sqr(Data[I] - MeanValue);

  N := Length(Data);
  if IsSample then
    Dec(N); // For sample standard deviation, use n-1 in the denominator

  Result := Sqrt(Sum / N);
end;




procedure TIWForm1.GerarCases;
var
  Vetor, Indices: array of Integer;
  Fatorial, Linha, I, j, k, kk, Temp: Integer;
  N: Integer;
  Matriz: array of array of Integer;
begin
  With TiwUserSession(WebApplication.Data) do
  begin
    // Calcular o fatorial de N (número total de permutações)
    Fatorial := 1;
    N := numcrit;
    for I := 2 to N do
      Fatorial := Fatorial * I;

    // Redimensionar a matriz para armazenar todas as permutações
    SetLength(Matriz, Fatorial, N);

    // Inicializar o vetor de entrada e índices
    SetLength(Vetor, N);
    for I := 0 to N - 1 do
      Vetor[I] := I + 1;

    SetLength(Indices, N);
    for I := 0 to N - 1 do
      Indices[I] := 0;

    // Copiar a primeira permutação
    Linha := 0;
    for I := 0 to N - 1 do
      Matriz[Linha][I] := Vetor[I];
    inc(Linha);

    // Algoritmo Heap's para gerar as permutações
    I := 0;
    while I < N do
    begin
      if Indices[I] < I then
      begin
        // Trocar elementos
        if I mod 2 = 0 then
          Temp := Vetor[0]
        else
          Temp := Vetor[Indices[I]];

        if I mod 2 = 0 then
          Vetor[0] := Vetor[I]
        else
          Vetor[Indices[I]] := Vetor[I];

        Vetor[I] := Temp;

        // Armazenar a nova permutação na matriz
        for j := 0 to N - 1 do
          Matriz[Linha][j] := Vetor[j];
        inc(Linha);

        // Incrementar o índice
        inc(Indices[I]);
        I := 0;
      end
      else
      begin
        Indices[I] := 0;
        inc(I);
      end;
    end;

   { if numcrit <= 7 then
    begin
      // Ordenar a matriz de permutações em ordem lexicográfica
      for I := 0 to Fatorial - 2 do
        for j := I + 1 to Fatorial - 1 do
        begin
          for k := 0 to N - 1 do
          begin
            if Matriz[I][k] < Matriz[j][k] then
              Break
            else if Matriz[I][k] > Matriz[j][k] then
            begin
              // Trocar as linhas
              for kk := 0 to N - 1 do
              begin
                Temp := Matriz[I][kk];
                Matriz[I][kk] := Matriz[j][kk];
                Matriz[j][kk] := Temp;
              end;
              Break;
            end;
          end;
        end;
    end; }

    SetLength(CasesOrdemcrit, Fatorial + 1, numcrit);
    SetLength(TiwUserSession(WebApplication.Data).NsolInst, Numalt);

    for I := 0 to Fatorial - 1 do
      for j := 0 to numcrit - 1 do
        TiwUserSession(WebApplication.Data).CasesOrdemcrit[I, j] := Matriz
          [I, j];
  end;
end;

procedure TIWForm1.IWLink1Click(Sender: TObject);
begin
  rgnmatcons.visible := true;
  rgnresultado.visible := false;
end;

procedure TIWForm1.IWLink2Click(Sender: TObject);
begin
  rgnestatistica.visible := false;
  rgnprimeiro.visible := true;
  IWLink2.visible := false;
  lbltexto.Caption :=
    'Percentage of cases in which the top positions defined the solution';
end;

procedure TIWForm1.IWLink3Click(Sender: TObject);
begin
  rgnestatistica.visible := true;
  rgnprimeiro.visible := false;
  IWLink2.visible := true;
  lbltexto.Caption := 'Simulation Statistical Parameters';

end;

procedure TIWForm1.IWLink4Click(Sender: TObject);
begin
  rgnmatcons.visible := true;
  rgnresultadoPromethee.visible := false;
end;

procedure TIWForm1.IWLink6Click(Sender: TObject);
begin
rgnregradecisaoPromethee.Visible := true;
rgnestatisticaPromethee.Visible := false;
end;

procedure TIWForm1.IWLink7Click(Sender: TObject);
begin
rgnregradecisao.Visible := true;
rgnestatistica.Visible := false;
end;

procedure TIWForm1.SetarCores;
Var
  I, j: Integer;
begin
  With TiwUserSession(WebApplication.Data) do
  begin
    SetLength(coresalt, 130);
    SetLength(coresalt1, TiwUserSession(WebApplication.Data).Numalt);
    coresalt[0] := $000000CC; // vermelho
    coresalt[1] := $0033BBFF; // amarelo
    coresalt[2] := $00CC9900; // azul
    coresalt[3] := $00BBBBFF; // rosa
    coresalt[4] := $005AE6C8; // verde
    coresalt[5] := $00FF00FF;
    coresalt[6] := $00808000;
    coresalt[7] := $00800000;
    coresalt[8] := $00000080;
    coresalt[9] := $0000FF00;
    coresalt[10] := $00008080;
    coresalt[11] := $00800080;
    coresalt[12] := $00FFFF00;
    coresalt[13] := $00C0DCC0;
    coresalt[14] := $00A4A0A0;
    coresalt[15] := $00E049BB;
    coresalt[16] := $005353FF;
    coresalt[17] := $000080FF;
    coresalt[18] := $00800040;
    coresalt[19] := $00434321;
    coresalt[20] := $00189AD8;
    coresalt[21] := $000000C1;
    coresalt[22] := $00DC0E65;
    coresalt[23] := $0000AE00;
    coresalt[24] := $002676CE;
    coresalt[25] := $008000FF; // índigo
    coresalt[26] := $00808080; // cinza claro
    coresalt[27] := $00FF4500; // laranja avermelhado
    coresalt[28] := $00FFD700; // ouro claro
    coresalt[29] := $00640000; // marrom escuro
    coresalt[30] := $008B4513; // marrom de sela
    coresalt[31] := $00A52A2A; // marrom de fogo
    coresalt[32] := $00D2691E; // chocolate
    coresalt[33] := $008BB000; // marrom escuro
    coresalt[34] := $00B22222; // tijolo
    coresalt[35] := $00FF6347; // tomate
    coresalt[36] := $00FF8C00; // laranja escuro
    coresalt[37] := $00FFA07A; // salmão claro
    coresalt[38] := $00FF7F50; // coral claro
    coresalt[39] := $00FF0000; // vermelho claro
    coresalt[40] := $00CD5C5C; // marrom claro
    coresalt[41] := $008B0000; // marrom escuro
    coresalt[42] := $00A52A2A; // marrom de fogo
    coresalt[43] := $00B22222; // tijolo
    coresalt[44] := $008B0000; // marrom escuro
    coresalt[45] := $00A52A2A; // marrom de fogo
    coresalt[46] := $00B22222; // tijolo
    coresalt[47] := $008B0000; // marrom escuro
    coresalt[48] := $00A52A2A; // marrom de fogo
    coresalt[49] := $00B22222; // tijolo
    coresalt[50] := $008B0000; // marrom escuro
    coresalt[51] := $00A52A2A; // marrom de fogo
    coresalt[52] := $00B22222; // tijolo
    coresalt[53] := $008B0000; // marrom escuro
    coresalt[54] := $00A52A2A; // marrom de fogo
    coresalt[55] := $00B22222; // tijolo
    coresalt[56] := $00FF69B4; // rosa quente
    coresalt[57] := $00FFE4E1; // rosa claro
    coresalt[58] := $00FF1493; // rosa profundo
    coresalt[59] := $00C71585; // vermelho violeta
    coresalt[60] := $00DB7093; // roxo claro
    coresalt[61] := $00DA70D6; // orquídea
    coresalt[62] := $009932CC; // índigo escuro
    coresalt[63] := $009ACD32; // verde amarelo
    coresalt[64] := $00C0FF3E; // verde limão brilhante
    coresalt[65] := $00AEEEEE; // azul celeste
    coresalt[66] := $000000CC; // vermelho
    coresalt[67] := $0033BBFF; // amarelo
    coresalt[68] := $00CC9900; // azul
    coresalt[69] := $00BBBBFF; // rosa
    coresalt[70] := $005AE6C8; // verde
    coresalt[71] := $00FF00FF;
    coresalt[72] := $00808000;
    coresalt[73] := $00800000;
    coresalt[74] := $00000080;
    coresalt[75] := $0000FF00;
    coresalt[76] := $00008080;
    coresalt[77] := $00800080;
    coresalt[78] := $00FFFF00;
    coresalt[79] := $00C0DCC0;
    coresalt[80] := $00A4A0A0;
    coresalt[81] := $00E049BB;
    coresalt[82] := $005353FF;
    coresalt[83] := $000080FF;
    coresalt[84] := $00800040;
    coresalt[85] := $00434321;
    coresalt[86] := $00189AD8;
    coresalt[87] := $000000C1;
    coresalt[88] := $00DC0E65;
    coresalt[89] := $0000AE00;
    coresalt[90] := $002676CE;
    coresalt[91] := $008000FF; // índigo
    coresalt[92] := $00808080; // cinza claro
    coresalt[93] := $00FF4500; // laranja avermelhado
    coresalt[94] := $00FFD700; // ouro claro
    coresalt[95] := $00640000; // marrom escuro
    coresalt[96] := $008B4513; // marrom de sela
    coresalt[97] := $00A52A2A; // marrom de fogo
    coresalt[98] := $00D2691E; // chocolate
    coresalt[99] := $008B0000; // marrom escuro
    coresalt[100] := $00B22222; // tijolo
    coresalt[101] := $00FF6347; // tomate
    coresalt[102] := $00FF8C00; // laranja escuro
    coresalt[103] := $00FFA07A; // salmão claro
    coresalt[104] := $00FF7F50; // coral claro
    coresalt[105] := $00FF0000; // vermelho claro
    coresalt[106] := $00CD5C5C; // marrom claro
    coresalt[107] := $008B0000; // marrom escuro
    coresalt[108] := $00A52A2A; // marrom de fogo
    coresalt[109] := $00B22222; // tijolo
    coresalt[110] := $008B0000; // marrom escuro
    coresalt[111] := $00A52A2A; // marrom de fogo
    coresalt[112] := $00B22222; // tijolo
    coresalt[113] := $008B0000; // marrom escuro
    coresalt[114] := $00A52A2A; // marrom de fogo
    coresalt[115] := $00B22222; // tijolo
    coresalt[116] := $008B0000; // marrom escuro
    coresalt[117] := $00A52A2A; // marrom de fogo
    coresalt[118] := $00B22222; // tijolo
    coresalt[119] := $008B0000; // marrom escuro
    coresalt[120] := $00A52A2A; // marrom de fogo
    coresalt[121] := $00B22222; // tijolo
    coresalt[122] := $00FF69B4; // rosa quente
    coresalt[123] := $00FFE4E1; // rosa claro
    coresalt[124] := $00FF1493; // rosa profundo
    coresalt[125] := $00C71585; // vermelho violeta
    coresalt[126] := $00DB7093; // roxo claro
    coresalt[127] := $00DA70D6; // orquídea
    coresalt[128] := $009932CC; // índigo escuro
    coresalt[129] := $009ACD32; // verde amarelo
  end;

end;

procedure TIWForm1.LimpaVariaveis;
begin
  With TIWUserSession(WebApplication.Data) do
  begin
  // Limpeza de arrays dinâmicos
  Finalize(NomeCrit);
  Finalize(NomeAlt);
  Finalize(TipoCrit);
  Finalize(Niveis);
  Finalize(MatrizConseq);
  Finalize(MatrizConseqNorm);
  Finalize(PesoCrit);
  Finalize(CasesOrdemCrit);
  Finalize(Max);
  Finalize(Min);
  Finalize(MatrizPOA);
  Finalize(MatrizPOAPromethee);
  Finalize(ResultadoROC);
  Finalize(ResultadoPromethee);
  Finalize(MatrizSol);
  Finalize(MatrizSolPromethee);
  Finalize(ResultAlt);
  Finalize(ResultAltPromethee);
  Finalize(ResultSol);
  Finalize(ResultSolPromethee);
  Finalize(coresalt);
  Finalize(coresalt1);
  Finalize(CaseSol);
  Finalize(CaseSolPromethee);
  Finalize(Crit10);
  Finalize(Crit12);
  Finalize(Crit13);
  Finalize(Crit14);
  Finalize(Crit20);
  Finalize(Crit21);
  Finalize(Crit23);
  Finalize(Crit24);
  Finalize(MatrizDifVG);
  Finalize(MatrizDifVGPromethee);
  Finalize(MediaDifSol);
  Finalize(MaxDifSol);
  Finalize(DesvioPadraoDifSol);
  Finalize(MinDifSol);
  Finalize(MediaDifSolPromethee);
  Finalize(MaxDifSolPromethee);
  Finalize(DesvioPadraoDifSolPromethee);
  Finalize(MinDifSolPromethee);
  Finalize(SobClassMatrix);
  Finalize(PositiveFlow);
  Finalize(NegativeFlow);
  Finalize(NetFlow);
  Finalize(matrizParaPar);
  Finalize(OrdemSol);
  Finalize(OrdemSolPromethee);
  Finalize(ordemalt);
  Finalize(ordemaltpromethee);
  Finalize(SimilaridadeAlt);
  Finalize(SomenteAditivo);
  Finalize(SomenteSobreclassificacao);
  Finalize(VetPOA);
  Finalize(MatrizPOA2);
  Finalize(Sconst);
  Finalize(VetPesos);
  Finalize(AltPOA);
  Finalize(PergNAdj);
  Finalize(MatrizAltSel);
  Finalize(VetNumPOA);
  Finalize(NumSolInst);
  Finalize(AltTamanhoPoa);
  Finalize(MatrizComparacaoAlt);
  Finalize(AmbasAltNaPOA);
  Finalize(NenhumaAltNaPOA);
  Finalize(CritName);
  Finalize(RestName);
  Finalize(AltName);
  Finalize(CritSigla);
  Finalize(AltSigla);
  Finalize(critcode);
  Finalize(altcode);
  Finalize(PrefDir);
  Finalize(vgmax);
  Finalize(vgmin);
  Finalize(Mediavgmax);
  Finalize(MediaVgMin);
  Finalize(MaxVgMax);
  Finalize(MaxVGMin);
  Finalize(MinVgMax);
  Finalize(MinVGMin);
  Finalize(AmplitudeVG);
  Finalize(difvg);
  Finalize(MediaAmplitudeVG);
  Finalize(MaxAmplitudeVG);
  Finalize(MinAmplitudeVG);
  Finalize(AmplitudeParAParinst);
  Finalize(ProbVitoriaParapar);
  Finalize(mediaprobvitoriaparapar);
  Finalize(maximoprobvitoriaparapar);
  Finalize(amplitudeparapar);
  Finalize(MediaGeralProbVitoriaAlt);
  Finalize(MaximoGeralProbVitoriaAlt);
  Finalize(MediaGeralProbDerrotaAlt);
  Finalize(MaximoGeralProbDerrotaAlt);
  Finalize(mediadifvg);
  Finalize(maximodifvg);
  Finalize(mediageralprobalt);

  // Reinicialização de variáveis escalares
  NumCrit := 0;
  Numalt := 0;
  excelverif := False;
  Total := 0.0;
  comp1 := 0.0;
  comp2 := 0.0;
  ind := 0.0;
  pref := 0.0;
  gauss := 0.0;
  SimilaridadeGeral := 0.0;
  MinMaxMaxNaoSol := 0.0;
  MinMaxMaxNaoSolPromethee := 0.0;
  NDA := 0;
  epsilon2 := 0.0;
  MinVK := 0.0;
  Epson := 0.0;


  end;
end;

procedure TIWForm1.PLEscolha(auxinst: Integer);
var
  i, j, k, Nrest, a, b, z, II, JJ, c: Integer; // Contadores
  NumCri, NumAlt, Numl, NAH, NumPOA: Integer; // Número de critérios,
  // num de alternativas, num de linhasna Matriz do Modelo de PL e número de restrições
  // provenientes de Avaliações holísticas, número de alternativas potencialmente ótimas
  MatA, MatCons: Array of Array of Double; // Matriz do Modelo de PL e Matriz de Consequências Normalizada
  VetPOA, Criorder: Array of Integer; // Vetor que armazena os índices das alternativas
  // potencialmente ótimas do problema, vetor que armazena as alternativas que passaram por uma
  // avaliação holística e vetor que armazena a ordenação dos critérios
  VetAH: Array of Array of Integer; // Matriz que armazena as informações das AH
  LimMin, LimMax: Array of Double; // Vetores que armazenam os limites inferior e
  // Superior da relação entre os critérios
  OPAH: Array of string; // Vetor que armazena o operador definido na relação estabelecida na AH
  G, ep: Double;
  // VARIÁVEIS LP SOLVE
  // ResultVector armazena os resultados obtidos pelo LP Solve (a função máxima e
  // coeficientes das variáveis de decisão).
  // FO recebe os coeficientes das variáveis de decisão na função objetivo e será
  // usado no LP Solve.
  // Limit recebe os valores das constantes das restrições e será usado no LP Solve.
  // X armazena
  NumDV, NRA, NRB: Integer; // número de var de decisão do modelo, inteiros auxiliares para a AH
  ResultVector: array of real;
  // typeprob determina se o problema é contínuo (typeprob = 0) ou binário (typeprob = 1).
  // No caso do fitradeoff, é sempre contínuo pois as constantes de escala são contínuas.
  // dirprob determina se o problema é de minimização (dirprob = 0) ou de maximização
  // (dirprob = 1), SF nos diz se houve solução viável.
  typeprob, dirprob, SF: Integer;
  FO, Op: array of string;
  limit: array of real;
  Rmod: array of AnsiString;
  var_int, var_bin: Array of Integer;
  /// /vetores que carregam var de decisão inteiras e binárias
  Lp: THandle;
  S1: string;
  MAXDI, MAXDJ: array [0 .. 2] of Double;
  AuxHEz: array of Integer; // vetor auxiliar para facilitar a elaboração de res_
  // trições com var binária da HE
  MSSG: Integer;
  MinVK: Double;
  pular, novo: boolean;
begin

  /// ///////// Recebendo o valor das variáveis salvas na usersession /////////////

  NumCri := TIWUserSession(WebApplication.Data).numcrit;
  NumAlt := TIWUserSession(WebApplication.Data).NumAlt;
  TIWUserSession(WebApplication.Data).NDA := NumAlt;
  NumPOA := TIWUserSession(WebApplication.Data).NDA;
  ep := TIWUserSession(WebApplication.Data).epsilon2;
  MinVK := TIWUserSession(WebApplication.Data).MinVK;
  /// Se inicia o número de linhas da Matriz como um valor suficiete para as restrições
  /// da PL, novas linhas serão acrescentadas de acordo ao número de AH realizadas///
  Numl := 2 * (NumCri - 1) + 4 + (NumPOA - 1);
  NumDV := NumCri;
  Setlength(MatCons, NumAlt, NumCri);
  Setlength(Criorder, NumCri);
  Setlength(MatA, Numl, NumCri + 1);
  Setlength(Op, Numl - 1);
  Setlength(limit, Numl - 1);
  Setlength(Rmod, Numl - 1);
  Setlength(VetPOA, NumPOA);
  Setlength(LimMax, NumCri - 1);
  Setlength(LimMin, NumCri - 1);
  Setlength(VetAH, NAH);
  Setlength(OPAH, NAH);
  Setlength(FO, NumCri);
  Setlength(ResultVector, NumCri + 1);

  typeprob := 0;
  /// Var de decisão são contínuas ///
  for i := 0 to NumAlt - 1 do
  begin
    for j := 0 to NumCri - 1 do
      MatCons[i, j] := TIWUserSession(WebApplication.Data)
        .MatrizConseqNorm[i, j];
  end;
  for i := 0 to NumCri - 1 do
    Criorder[i] := i;
  for i := 0 to NumCri - 2 do
  begin
    // setlength(LimMin, i + 1);
    LimMin[i] := 0;
    LimMax[i] := 1;
  end;

  for i := 0 to NumPOA - 1 do
  begin
    VetPOA[i] := i;
  end;

  /// ///////////////////////////////////////////////////////////////////////////
  /// Limpando a Matriz do problema de possíveis valores provindos de sua inicialização ///
  for i := 0 to Numl - 1 do
    for j := 0 to NumCri do
      MatA[i, j] := 0;

  /// //////////// Programação Linear - MODELO PROB DE ESCOLHA //////////////////

  /// Preenchendo as Restrições referentes ao limite superior da relação dos critérios ///
  /// transformando-a em uma restrição do tipo '<=' para seguir o padrão adotado no SAD ///
  for i := 0 to NumCri - 2 do
  begin
    MatA[i + 1, i + 1] := (-1) * LimMax[i];
    MatA[i + 1, i + 2] := 1;
    Op[i] := '<=';
    INC(Nrest);
  end;

  a := Nrest;

  /// Preenchendo as Restrições referentes ao limite inferior da relação dos critérios ///
  /// Essa restrição já é do tipo '<=' que segue o padrão adotado no SAD ///
  for i := a to a + NumCri - 2 do
  begin
    MatA[i + 1, i - a + 1] := LimMin[i - a];
    MatA[i + 1, i - a + 2] := (-1) * 1;
    Op[i] := '<=';
    INC(Nrest);
  end;

  a := Nrest;

  /// / Adiciona variáveis binárias ao problema caso necessário ////////////////
  Setlength(FO, NumDV);
  Setlength(var_bin, (NumDV - NumCri));
  Setlength(ResultVector, NumDV + 1);
  for j := 0 to length(var_bin) - 1 do
    var_bin[j] := NumCri + j;
  /// //////////////////////////////////////////////////////////////////////////
  a := Nrest;

  a := Nrest;

  /// Preenchendo aa nova restrição Kn > limiar///
  MatA[a + 1, NumCri] := -1;
  MatA[a + 1, 0] := -MinVK;
  Op[a] := '<=';
  INC(Nrest);

  a := Nrest;

  /// Preenchendo as Restrições referentes à Normalização dos Pesos///
  for j := 0 to NumCri - 1 do
    MatA[a + 1, j + 1] := 1;
  MatA[a + 1, 0] := 1;
  Op[a] := '=';
  INC(Nrest);

  a := Nrest;

  a := Nrest;
  NumPOA := 0;


  for z := 0 to TIWUserSession(WebApplication.Data).NDA - 1 do
  begin
    k := VetPOA[z];
    b := a;

    /// Preenchendo as Restrições de Potencial Otimalidade ///
    /// transformando-a em uma restrição do tipo '<=' para seguir o padrão adotado no SAD ///
    for i := a to a + TIWUserSession(WebApplication.Data).NDA - 1 do
    begin
      if VetPOA[i - a] <> k then
      begin
        for j := 0 to NumCri - 1 do
          MatA[b + 1, j + 1] := MatCons[VetPOA[i - a], Criorder[j]] - MatCons
            [k, Criorder[j]];
        Op[b] := '<=';
        MatA[b + 1, 0] := -TIWUserSession(WebApplication.Data).Epson;
        INC(Nrest);
        INC(b);
      end;
    end;

    /// Preenchendo a função objetivo, onde as var de decisão são as constantes de ///
    /// escala e os coeficientes são as consequências da alternativa normalizada ///
    for i := 0 to NumCri - 1 do
    begin
      MatA[0, i + 1] := MatCons[k, Criorder[i]];
      FO[i] := TrocaVP(MatCons[k, Criorder[i]]);
    end;
    /// / Caso hajam var binárias no problema, seus coef na FO devem ser 0 ////
    for i := NumCri to NumDV - 1 do
    begin
      MatA[0, i + 1] := 0;
      FO[i] := '0';
    end;

    /// Preenchendo o lado direito da restrição no vetor utilizado no LP_Solve ///
    for i := 0 to Numl - 2 do
      limit[i] := MatA[i + 1, 0];

    /// Preenchendo as restrições no vetor utilizado no LP_Solve ///
    for i := 0 to Numl - 2 do
    begin
      Rmod[i] := TrocaVP(MatA[i + 1, 1]);
      for j := 2 to NumDV do
        Rmod[i] := Rmod[i] + ' ' + TrocaVP(MatA[i + 1, j]);
    end;

    /// Problema de maximização ///
    dirprob := 1;

    /// Resolvendo o PPL ///
    SolveLP(NumDV, Nrest, dirprob, typeprob, FO, Op, limit, var_int, var_bin,
      Rmod, ResultVector, Lp, MSSG);

    TIWUserSession(WebApplication.Data).Vgmax[auxinst, z] := Resultvector[0];

    SF := MSSG;
    if SF = 0 then
    begin
      /// Verificando se há algum K não nulo. ///
      Setlength(TIWUserSession(WebApplication.Data).Sconst, NumCri);
      b := 0;
      for i := 1 to NumCri do
        if ResultVector[i] <> 0 then
          break
        else
          INC(b);

      if b <> NumCri then
      begin
        /// Atualização do Vetor de Alternativas Potencialmente Ótimas ///
        INC(NumPOA);
        Setlength(TIWUserSession(WebApplication.Data).VetPOA, NumPOA);
        Setlength(TIWUserSession(WebApplication.Data).VetPesos, NumPOA, NumCri);

        TIWUserSession(WebApplication.Data).VetPOA[NumPOA - 1] := k;
         TIWUserSession(WebApplication.Data).VGMax[auxinst,NumPOA - 1] := ResultVector
          [0];
        for i := 0 to NumCri - 1 do
        begin
          TIWUserSession(WebApplication.Data).Sconst[i] := ResultVector[i + 1];
          TIWUserSession(WebApplication.Data).VetPesos[NumPOA - 1, i] :=
            ResultVector[i + 1];
        end;
      end;

        // ** Só é preciso solucionar o problema de minimização <=> a alternativa é POA **//
        // ** Além disso, as restrições e a FO é a mesma **//
        /// Problema de maximização ///
         dirprob := 0;
          Setlength(ResultVector, 0);
          Setlength(ResultVector, NumDV + 1);
          /// Resolvendo o PPL ///
          SolveLP(NumDV, Nrest, dirprob, typeprob, FO, Op, limit, var_int,
          var_bin, Rmod, ResultVector, Lp, MSSG);
          if MSSG = 0 then
          TIWUserSession(WebApplication.Data).VGMin[auxinst, z] := ResultVector
          [0];


    end;

    /// Limpando as restrições de Otimalidade da Matriz.
    for j := 0 to NumCri do
    begin
      MatA[0, j] := 0;
      for i := a + 1 to Nrest do
        MatA[i, j] := 0;
    end;
    Nrest := a;
    Setlength(ResultVector, 0);
    Setlength(ResultVector, NumDV + 1);
  end;

  // ** Caso nenhuma POA seja encontrada, o problema pode estar no epson **//
  if NumPOA = 0 then
  begin
    TIWUserSession(WebApplication.Data).Epson := 0;
    PLEscolha(auxinst);
    exit;
  end;

  TIWUserSession(WebApplication.Data).NDA := NumPOA;

  inc(TIWUserSession(WebApplication.Data).nsolinst[NumPOA]);



  /// ///////////////////////////////////////////////////////////////////////////
  /// //////////////////////Max diferença////////////////////////////////////////
  if (NumPOA > 1) and (NumPOA <= 3) and (ep <> 0) then
  begin

    /// Preenchendo o lado direito da restrição no vetor utilizado no LP_Solve //
    Setlength(limit, Nrest);
    for i := 0 to Nrest - 1 do
      limit[i] := MatA[i + 1, 0];

    /// Preenchendo as restrições no vetor utilizado no LP_Solve ///
    Setlength(Rmod, Nrest);
    for i := 0 to Nrest - 1 do
    begin
      Rmod[i] := TrocaVP(MatA[i + 1, 1]);
      for j := 2 to NumDV do
        Rmod[i] := Rmod[i] + ' ' + TrocaVP(MatA[i + 1, j]);
    end;
    Setlength(Op, Nrest);
    a := 1;

    c := 0;
    for II := 0 to NumPOA - 1 do
      for JJ := II + 1 to NumPOA - 1 do
      begin
        /// ///////// Verifica a max diferença entre II - JJ////////////////////

        /// / Preenche a função objetivo ///////////////////////////////////////
        for k := 0 to NumCri - 1 do
          FO[k] := TrocaVP
            (MatCons[VetPOA[II], Criorder[k]] - MatCons[VetPOA[JJ], Criorder[k]]
            );
        /// / Caso hajam var binárias no problema, seus coef na FO devem ser 0 ////
        for i := NumCri to NumDV - 1 do
        begin
          FO[i] := '0';
        end;

        /// Resolvendo o PPL ///
        SolveLP(NumDV, Nrest, dirprob, typeprob, FO, Op, limit, var_int,
          var_bin, Rmod, ResultVector, Lp, MSSG);

        SF := MSSG;

        if SF = 0 then
        begin
          /// Verificando se há algum K não nulo. ///
          b := 0;
          for i := 1 to NumCri do
            if ResultVector[i] <> 0 then
              break
            else
              INC(b);

          MAXDI[II] := ResultVector[0];
        end; // If SF

        /// ///////// Verifica a max diferença entre JJ - II////////////////////
        /// / Preenche a função objetivo ///////////////////////////////////////
        for k := 0 to NumCri - 1 do
          FO[k] := TrocaVP
            (MatCons[VetPOA[JJ], Criorder[k]] - MatCons[VetPOA[II], Criorder[k]]
            );
        /// / Caso hajam var binárias no problema, seus coef na FO devem ser 0 ////
        for i := NumCri to NumDV - 1 do
        begin
          FO[i] := '0';
        end;

        /// Resolvendo o PPL ///
        SolveLP(NumDV, Nrest, dirprob, typeprob, FO, Op, limit, var_int,
          var_bin, Rmod, ResultVector, Lp, MSSG);

        SF := MSSG;
        if SF = 0 then
        begin
          /// Verificando se há algum K não nulo. ///
          b := 0;
          for i := 1 to NumCri do
            if ResultVector[i] <> 0 then
              break
            else
              INC(b);

          MAXDJ[II] := ResultVector[0];
        end; // If SF
        {
          for j := 0 to numalt - 1 do
          TIWUserSession(WebApplication.Data).MatrizPOA2[auxinst, i] := -1;

          for I := 0 to length(vetPOA) - 1 do
          begin
          TIWUserSession(WebApplication.Data).MatrizPOA2[auxinst, i] := vetPOA[i];

          TIWUserSession(WebApplication.Data).AltPOA[vetpoa[i]] := 1 + TIWUserSession(WebApplication.Data).AltPOA[vetpoa[i]];
          end;
          }
        /// / Deve-se considerar o módulo da máx diferença entre as alternativas ////�
        if MAXDI[II] < 0 then
          MAXDI[II] := -MAXDI[II];
        if MAXDJ[II] < 0 then
          MAXDJ[II] := -MAXDJ[II];

        if (MAXDI[II] <= ep) and (MAXDJ[II] <= ep) then
          INC(c);

      end; // for JJ

  end; // if numpoa
  //delete_lp(lp);

  //free_lp(lp);

  DSiTrimWorkingSet;
end;

procedure TIWForm1.DSiTrimWorkingSet;
var
  hProcess: THandle;
begin
  hProcess := OpenProcess(PROCESS_SET_QUOTA, false, GetCurrentProcessId);
  try
    SetProcessWorkingSetSize(hProcess, $FFFFFFFF, $FFFFFFFF);
  finally CloseHandle(hProcess); end;
end; { DSiTrimWorkingSet }



procedure TIWForm1.btnshowresultsClick(Sender: TObject);
var
  I, j, k, inst: Integer;
  bar, barpromethee: array of TBarSeries;
  Fatorial: Double;
  Naltsel: Integer;
  soma: Double;
  aux1, aux2: string;
  valor1, valor2: Integer;
begin
  With TiwUserSession(WebApplication.Data) do
  begin
    LimpaVariaveis;


    aux1 := edtnumalt.Text;
    aux2 := cmbbxcriteria.Items[cmbbxcriteria.ItemIndex];
    if (TryStrToInt(aux1, valor1) = false) or
      ((TryStrToInt(aux2, valor2) = false)) then
    begin
      WebApplication.ShowMessage('Fill in the fields correctly!');
      exit;
    end;

    //GerarCases;
    Numalt := StrToInt(edtnumalt.Text);
    numcrit := StrToInt(cmbbxcriteria.Items[cmbbxcriteria.ItemIndex]);
    SetLength(Nomecrit, numcrit);
    SetLength(Nomealt, Numalt);
    SetLength(tipocrit, numcrit);
    SetLength(niveis, numcrit);
    SetLength(MatrizConseq, Numalt, numcrit);
    SetLength(MatrizConseqNorm, Numalt, numcrit);
    SetLength(Max, numcrit);
    SetLength(Min, numcrit);
    epsilon2 := 0.01;
    Epson := 0.01;
    MinVK := 0.000000001;
    Setlength(TIWUserSession(WebApplication.Data).MatrizPOA2, length
        (CasesOrdemcrit), NumAlt);
    Setlength(TIWUserSession(WebApplication.Data).MatrizAltSel, length
        (CasesOrdemcrit), NumAlt);
    Setlength(TIWUserSession(WebApplication.Data).AltPOA, NumAlt);
    Setlength(TIWUserSession(WebApplication.Data).matrizsol, 1, NumAlt);
    SetLength(TIWUserSession(WebApplication.Data).ResultSol, 1);
    Setlength(TIWUserSession(WebApplication.Data).VetNumPoa, 1);
    SetLength(TIWUserSession(WebApplication.Data).NumSolInst, Numalt);
    SetLength(TIWUserSession(WebApplication.Data).AltTamanhoPOA, Numalt, Numalt-1);
    SetLength(TIWUserSession(WebApplication.Data).MatrizComparacaoAlt, Numalt, Numalt);
    SetLength(TIWUserSession(WebApplication.Data).AmbasAltNaPOA, Numalt, Numalt);
    SetLength(TIWUserSession(WebApplication.Data).NenhumaAltNaPOA, Numalt, Numalt);
    Setlength(TIWUserSession(WebApplication.Data).VGMax, 1,TIWUserSession(WebApplication.Data).Numalt);
    Setlength(TIWUserSession(WebApplication.Data).VGMin, 1,TIWUserSession(WebApplication.Data).Numalt);
    setlength(nsolinst, numalt);
    Ninst := 10000;
    for j := 0 to Numalt - 1 do
    begin
      Nomealt[j] := 'Alt ' + inttostr(j + 1);
    end;
    for I := 0 to numcrit - 1 do
    begin
      TiwUserSession(WebApplication.Data).tipocrit[I] := 1;
      Nomecrit[I] := 'Crit ' + inttostr(I + 1);
    end;

    for inst := 0 to Ninst - 1 do
    begin
      for j := 0 to Numalt - 1 do
        for I := 0 to numcrit - 1 do
        begin
          TiwUserSession(WebApplication.Data).MatrizConseq[j, I] := Random;
          MatrizConseqNorm[j, I] := MatrizConseq[j, I];
        end;


    epsilon2 := 0.01;
    Epson := 0.01;
    MinVK := 0.000000001;
    ResultSol[0] := 1;

    PlEscolha(0);

    end;

      chrtResultado.RemoveAllSeries;
      // chrtresultado.ClearChart;
      With TiwUserSession(WebApplication.Data) do
      begin
        j := 0;
        SetLength(bar, j + 1);
        bar[j] := TBarSeries.Create(self);
        chrtResultado.AddSeries(bar[j]);
        for I := 0 to Length(NsolInst) - 1 do
        begin
          if Round(TiwUserSession(WebApplication.Data).NsolInst[I]) <> 0 then
          begin
            bar[0].AddXY
              (j * 2, (NsolInst[I] / TiwUserSession(WebApplication.Data).Ninst)
                * 100, FormatFloat('0.00', (NsolInst[I] / Ninst) * 100) + '%');
            chrtResultado.BottomAxis.Items.Add(j * 2, inttostr(I + 1));
            // chrtResultado.Legend.Item[0].Text := inttostr(i + 1) + ' Solutions ';
            inc(j);
          end;
        end;
      end;

      chrtResultado.LeftAxis.Increment := 0.2; // ?
      // chrtResultado.BottomAxis.visible := false;
      chrtResultado.Legend.visible := false;
      chrtResultado.BottomAxis.LabelsAlternate;
      chrtResultado.BottomAxis.Items.Automatic := false;
      chrtResultado.LeftAxis.Maximum := 100;
      chrtResultado.LeftAxis.MaximumOffset := 16;
      chrtResultado.LeftAxis.Minimum := 0;
      // chrtresultado.LeftAxis.AutomaticMaximum := false;
      chrtResultado.LeftAxis.Increment := 20;
      chrtResultado.Title.Caption :=
        'Percentage of cases with X number of solutions';
      imgResultado.Picture.Bitmap.Assign(chrtResultado.TeeCreateBitmap
          (clWebPURPLE, rect(0, 0, imgResultado.Width, imgResultado.Height)));

      QtdMedia := 0;
      j := 0;
      for I := 0 to Length(NsolInst) - 1 do
        if NsolInst[I] <> 0 then
        begin
          TiwUserSession(WebApplication.Data).QtdMedia := QtdMedia + (I + 1)
            * NsolInst[I] / Ninst;
          inc(j);
        end;
      lblmedia.Caption := FormatFloat('0.00', QtdMedia);

      mediana := 0;
      j := 0;
      for I := 0 to Length(NsolInst) - 1 do
        if NsolInst[I] <> 0 then
        begin
          j := j + Round(NsolInst[I]);
          if j = Round(Ninst / 2) then
          begin
            mediana := ((I + 1) + (I + 2)) / 2;
            Break;
          end;
          if j > Round(Ninst / 2) then
          begin
            mediana := I + 1;
            Break;
          end;
        end;
      lblmediana.Caption := FloatToStr(mediana);

      moda := 0;
      j := 0;
      for I := 0 to Length(NsolInst) - 1 do
        if NsolInst[I] <> 0 then
        begin
          if j < NsolInst[I] then
          begin
            j := Round(NsolInst[I]);
            moda := I + 1;
          end;

        end;

      lblmoda.Caption := FloatToStr(moda);

      DesvioPadrao := 0;
      for I := 0 to Length(NsolInst) - 1 do
        if NsolInst[I] <> 0 then
        begin
          TiwUserSession(WebApplication.Data).DesvioPadrao :=
            DesvioPadrao + NsolInst[I] * Sqr(abs((I + 1) - QtdMedia))
        end;
      DesvioPadrao := Sqrt(DesvioPadrao / Ninst);

      lbldesviopadrao.Caption := FormatFloat('0.00', DesvioPadrao);


    SetarCores;


      grdnsolutions.RowCount := 2;
      grdnsolutions.ColumnCount := 2;

      for j := 0 to grdnsolutions.ColumnCount - 1 do
        grdnsolutions.Cell[0, j].Font.Style := [fsbold];

      // for i := 0 to grdresultados.RowCount - 1 do
      // grdnsolutions.Cell[i,0].Font.Style := [fsbold];

      grdnsolutions.Cell[0, 0].Text := 'Number of Solutions';
      grdnsolutions.Cell[0, 1].Text := 'Percentage';
      j := 0;
      for I := 0 to Length(NsolInst) - 1 do
      begin
        if Round(NsolInst[I]) = 0 then
          continue
        else
        begin
          grdnsolutions.RowCount := 2 + j;
          if I = 0 then
            grdnsolutions.Cell[1 + j, 0].Text := '1 Solution'
          else
            grdnsolutions.Cell[j + 1, 0].Text := inttostr(I + 1) + ' Solutions';
          grdnsolutions.Cell[j + 1, 1].Text := FormatFloat
            ('0.00', (NsolInst[I] / Ninst) * 100) + '%';
          inc(j);
        end;
      end;
      for j := 0 to grdnsolutions.ColumnCount - 1 do
        for I := 0 to grdnsolutions.RowCount - 1 do
        begin
          grdnsolutions.Cell[I, j].Alignment := taCenter;
          grdnsolutions.Cell[I, j].wrap := true;
          // grdresultados.Cell[0,j].Font.Style := [fsbold];
        end;
      rgnmatcons.visible := false;
      rgnresultado.visible := true;
      // btnshowresults.Enabled := true;

    end;

end;

procedure TIWForm1.cmbbxauxiliarAsyncChange(Sender: TObject;
  EventParams: TStringList);
var
  I: Integer;
begin
  With TiwUserSession(WebApplication.Data) do
  begin
    for I := 0 to numcrit - 1 do
    begin
      if (cmbbxTipoCrit[I].ItemIndex = 2) or (cmbbxTipoCrit[I].ItemIndex = 3)
        then
        cmbbxNiveisCrit[I].visible := true
      else
        cmbbxNiveisCrit[I].visible := false;
    end;
  end;
end;

procedure TIWForm1.cmbbxauxiliarChange(Sender: TObject);
var
  I: Integer;
begin
  With TiwUserSession(WebApplication.Data) do
  begin
    for I := 0 to numcrit - 1 do
    begin
      if (cmbbxTipoCrit[I].ItemIndex = 2) or (cmbbxTipoCrit[I].ItemIndex = 3)
        then
        cmbbxNiveisCrit[I].visible := true
      else
        cmbbxNiveisCrit[I].visible := false;
    end;
  end;
end;

procedure TIWForm1.rgnmatconsCreate(Sender: TObject);
begin
  rgnmatcons.Align := AlClient;
end;

procedure TIWForm1.rgnresultadoCreate(Sender: TObject);
begin
  rgnresultado.Align := AlClient;
end;

procedure TIWForm1.rgnresultadoPrometheeCreate(Sender: TObject);
begin
rgnresultadoPromethee.Align := Alclient;
end;

initialization

TIWForm1.SetAsMainForm;

end.
