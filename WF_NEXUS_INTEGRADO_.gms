$Title WF-Nexus Aguacate - Modelo E1 y sensibilidad integrada

* Modelo mensual para los seis municipios.
* Se usa la formulacion con area fija y cuatro etapas.
* Cada tecnologia se resuelve por separado.
*
* E1 es el caso de referencia. La ETc se toma del Excel de entrada.
* A1, A2: limites de entrega mensual de agua.
* R80, R60, R40, R20: reduccion de limites anuales REPDA.
* H0, H1, H2, H3: escenarios de captacion y almacenamiento.
* Ky 1.10, 1.20 y 1.30 se aplican despues de la optimizacion.

$if not set INPUT  $set INPUT  BASE_WF_NEXUS_E1_PRINCIPAL.xlsx
$if not set OUTPUT $set OUTPUT RESULTADOS_WF_NEXUS_E1_INTEGRADO.xlsx
$if not set INPUTE2 $set INPUTE2 BASE_GAMS_KC_E2.xlsx

option limrow=0, limcol=0, solprint=off, optcr=0;

Sets
    r Municipios /
        Uruapan, Tancitaro, Periban, Ario,
        SalvadorEscalante, Tacambaro /
    m Meses /
        Ene, Feb, Mar, Abr, May, Jun,
        Jul, Ago, Sep, Oct, Nov, Dic /
    k Tecnologias /
        Aspersion70, Microaspersion80, Goteo95 /
    y ValoresKy /
        Ky110, Ky120, Ky130 /
    s Corridas /
        E1, A1, A2, R80, R60, R40, R20,
        H0, H1, H2, H3, E2 /
    lt Fuentes /
        LimSub_m3, LimSup_m3 /
    cc Clima /
        P_mm, Pef_mm /
    hc Historico /
        AreaProd_ha, Prod_ton /
    md IndicadoresMensuales /
        Demanda_m3, Subterranea_m3, Superficial_m3,
        LluviaUsada_m3, Deficit_m3, Almacen_m3,
        Cobertura_pct, ETm_m3, ETaEstimada_m3 /
    mu IndicadoresMunicipales /
        Area_ha, ProduccionReferencia_ton, Demanda_m3,
        Subterranea_m3, Superficial_m3, LluviaCaptada_m3,
        LluviaUsada_m3, Deficit_m3, Cobertura_pct,
        PresionSub_pct, PresionSup_pct, MaxDefRelMensual_pct,
        ETm_m3, ETaEstimada_m3, EstresET_pct /
    rg IndicadoresRegionales /
        Area_ha, ProduccionReferencia_ton, Demanda_m3,
        Subterranea_m3, Superficial_m3, LluviaCaptada_m3,
        LluviaUsada_m3, Deficit_m3, Cobertura_pct,
        MaxDefRelMensual_pct, ETm_m3, ETaEstimada_m3,
        EstresET_pct /
    pk IndicadoresProductivos /
        ProduccionReferencia_ton, ProduccionAjustada_ton,
        PerdidaProduccion_ton, FactorRendimiento_pct,
        EstresET_pct, Ky, CoberturaRiego_pct /
    au Comprobaciones /
        Obj1Def_m3, Obj2Sub_m3, Obj3Sup_m3, Obj4DefMax_pct,
        ModelStat1, ModelStat2, ModelStat3, ModelStat4,
        SolveStat1, SolveStat2, SolveStat3, SolveStat4,
        BalanceMax_m3, AlmacenMax_m3,
        DeltaDefMunicipal_m3, DeltaSubMunicipal_m3,
        DeltaSupMunicipal_m3, DeltaDefRegional_m3,
        DeltaSubRegional_m3, DeltaSupRegional_m3,
        DiferenciaObjetivo4_pct /
    rev ComparacionE1 /
        Demanda_m3, Deficit_m3, Cobertura_pct,
        DeltaDeficit_m3, DeltaCobertura_pp, CoberturaCompleta,
        DeficitRelMax_pct, EstresET_pct /
;
Alias(k,kk);

Parameters
    eta(k)              Eficiencia de aplicacion
    kyValue(y)          Factor de respuesta del rendimiento
    dias(m)             Dias por mes del calendario de 365 dias
    Vreq70(r,m)         Requerimiento antiguo para control
    Lim(r,lt)           Derechos ajustados de referencia anuales m3
    LimActivo(r,lt)     Limites aplicables en la corrida
    Clima(r,m,cc)       Precipitacion total y efectiva mm
    AreaMax(r)          Superficie bajo riego ha
    Hist(r,hc)          Datos de produccion SIAP
    EToMensual(r,m)     ETo AgERA5 mm por mes
    KcEfectivo(r,m)     Kc E1 mensual ponderado por ETo diaria
    ETcArchivo(r,m)     ETc mensual de E1 en mm
    ETcE2(r,m)          ETc mensual de E2 si existe su archivo
    ETcActiva(r,m)     ETc del escenario actual
    ReqActive(r,m)      Requerimiento bruto mensual m3 por ha
    CaptPotHa(r,m)      Captacion potencial m3 por ha
    CapAlmHa(r)         Capacidad local de almacenamiento m3 por ha
    LimSubMes(r,m)      Limite mensual de agua subterranea m3
    LimSupMes(r,m)      Limite mensual de agua superficial m3
    DefMunStar(r)       Deficit anual municipal de la etapa 3
    SubMunStar(r)       Uso subterraneo anual municipal de etapa 3
    SupMunStar(r)       Uso superficial anual municipal de etapa 3
    ETmAnual(r)         Evapotranspiracion potencial del cultivo m3
    ETaEstimada(r)      Evapotranspiracion estimada con deficit m3
    EstresET(r)         Fraccion de ETm no satisfecha
    GestionMensual(r,m,s,k,md)
    GestionMunicipal(r,s,k,mu)
    GestionRegional(s,k,rg)
    ProduccionKyMunicipal(r,s,k,y,pk)
    ProduccionKyRegional(s,k,y,pk)
    AuditoriaEtapas(s,k,au)
    ResumenComparativo(r,s,k,rev)
    ParametrosEscenario(s,*)
    EscenariosEjecutados(s) Escenario incluido 1 si 0 no
;

Set sRun(s) Corridas con datos disponibles;
sRun(s) = yes;
sRun('E2') = no;

eta('Aspersion70')      = 0.70;
eta('Microaspersion80') = 0.80;
eta('Goteo95')          = 0.95;

kyValue('Ky110') = 1.10;
kyValue('Ky120') = 1.20;
kyValue('Ky130') = 1.30;

dias('Ene')=31; dias('Feb')=28; dias('Mar')=31;
dias('Abr')=30; dias('May')=31; dias('Jun')=30;
dias('Jul')=31; dias('Ago')=31; dias('Sep')=30;
dias('Oct')=31; dias('Nov')=30; dias('Dic')=31;

Scalars
    fracCaptura      Fraccion de lluvia no efectiva capturable
    fracAlmacen      Fraccion de captura anual admisible como almacenamiento
    factorLim        Fraccion de los limites anuales REPDA
    rho             Factor de concentracion mensual de extracciones
    usaTopeMensual  Indicador para A1 y A2
    epsDef          /0.001/
    epsSub          /0.001/
    epsDef4         /10/
    epsSub4         /10/
    epsSup4         /10/
    DefStar
    SubStar
    SupStar
    MaxBalance
    MaxBalanceAlm
    MaxDefMun
    MaxSubMun
    MaxSupMun
    MaxDefReg
    MaxSubReg
    MaxSupReg
    MaxRelReal
    GIR70Calculado
    GIR70Esperado    /308225164.3230495/
    GIR70Anterior
    GIR70E2Esperado  /304509838.3308986/
    GIR70E2Calculado
;

* Los datos de E1 deben ser los mismos que usamos en el ultimo modelo.
$if not exist "%INPUT%" $abort "No se encontro el Excel principal de E1. Revisar --INPUT."

$onecho > lectura_e1_ky.txt
par=Vreq70        rng=ENTRADA_GAMS!B2:D73   rdim=2 cdim=0
par=Lim           rng=ENTRADA_GAMS!F1:H7    rdim=1 cdim=1
par=Clima         rng=ENTRADA_GAMS!J1:M73   rdim=2 cdim=1
par=AreaMax       rng=ENTRADA_GAMS!O2:P7    rdim=1 cdim=0
par=Hist          rng=ENTRADA_GAMS!R1:T7    rdim=1 cdim=1
par=EToMensual    rng=ETO_GAMS!A2:C73       rdim=2 cdim=0
par=KcEfectivo    rng=KC_EFECTIVO!A2:C73    rdim=2 cdim=0
par=ETcArchivo    rng=ETC_GAMS!A2:C73       rdim=2 cdim=0
$offecho

$call gdxxrw "%INPUT%" output=entrada_e1_ky.gdx @lectura_e1_ky.txt trace=3
$if errorlevel 1 $abort "Error leyendo hojas de E1. Comprobar rangos y nombres."
$gdxin entrada_e1_ky.gdx
$load Vreq70 Lim Clima AreaMax Hist EToMensual KcEfectivo ETcArchivo
$gdxin

abort$(smin(r,AreaMax(r)) <= 0) "Area de riego no valida";
abort$(smin(r,Hist(r,'AreaProd_ha')) <= 0) "Area cosechada no valida";
abort$(smin(r,Hist(r,'Prod_ton')) <= 0) "Produccion historica no valida";
abort$(smin((r,m),EToMensual(r,m)) <= 0) "ET0 mensual no valida";
abort$(smin((r,m),KcEfectivo(r,m)) < 0.599 or
       smax((r,m),KcEfectivo(r,m)) > 0.851)
    "Kc efectivo E1 fuera de los valores establecidos";
abort$(smax((r,m),abs(ETcArchivo(r,m)-EToMensual(r,m)*KcEfectivo(r,m)))>1e-6)
    "ETc de E1 no coincide con ETo y Kc";

GIR70Calculado = sum((r,m),
    10*AreaMax(r)*max(0,ETcArchivo(r,m)-Clima(r,m,'Pef_mm'))/0.70);
GIR70Anterior = sum((r,m),Vreq70(r,m));
abort$(abs(GIR70Calculado-GIR70Esperado)>50)
    "El archivo no reproduce la demanda de referencia E1";
abort$(abs(GIR70Calculado-GIR70Anterior)<1000000)
    "Parece que se cargo la demanda anterior en lugar de la de E1";

* La comparacion E2 se activa solo si se encuentra el archivo que ya
* habiamos preparado para la floracion marcena (15 de marzo).
* Se importa su ETc, sin reconstruirla con otro metodo.
* Si no hay archivo E2, el valor de control queda en cero.
GIR70E2Calculado=0;
$ifthen exist "%INPUTE2%"
$onecho > lectura_e2_ky.txt
par=ETcE2 rng=ETC_GAMS!A2:C73 rdim=2 cdim=0
$offecho
$call gdxxrw "%INPUTE2%" output=entrada_e2_ky.gdx @lectura_e2_ky.txt trace=3
$if errorlevel 1 $abort "No se pudo leer ETc de E2"
$gdxin entrada_e2_ky.gdx
$load ETcE2
$gdxin
abort$(smin((r,m),ETcE2(r,m)) <= 0) "ETc de E2 incompleta";
GIR70E2Calculado = sum((r,m),
    10*AreaMax(r)*max(0,ETcE2(r,m)-Clima(r,m,'Pef_mm'))/0.70);
abort$(abs(GIR70E2Calculado-GIR70E2Esperado)>50)
    "La demanda de E2 no coincide con el escenario anterior";
sRun('E2')=yes;
$else
* Si falta E2, el parametro debe tener una asignacion explicita.
* E2 no se ejecuta: estos valores sirven solo para compilar.
ETcE2(r,m)=ETcArchivo(r,m);
$endif

EscenariosEjecutados(s)=1$sRun(s);

* Variables hidricas del modelo original.
Positive Variables
    Area(r)       Superficie irrigada
    WSub(r,m)     Agua subterranea
    WSup(r,m)     Agua superficial
    Capt(r,m)     Agua de lluvia captada
    WAlm(r,m)     Agua de lluvia utilizada
    Store(r,m)    Inventario de lluvia almacenada
    Def(r,m)      Demanda no cubierta
;
Variables
    DefTotal
    GWTotal
    SurfaceTotal
    MaxDefRel
;

Area.fx(r)=AreaMax(r);
MaxDefRel.lo=0;
MaxDefRel.up=1;

Equations
    DefDeficit
    DefGW
    DefSurface
    BalanceAgua(r,m)
    LimiteSub(r)
    LimiteSup(r)
    LimiteSubMensual(r,m)
    LimiteSupMensual(r,m)
    LimiteCapt(r,m)
    BalanceStoreEne(r)
    BalanceStore(r,m)
    LimiteStore(r,m)
    FijarDeficit
    FijarSub
    FijarDeficit4
    FijarSub4
    FijarSup4
    FijarDeficitLo
    FijarSubLo
    FijarSupLo
    FijarDefMun(r)
    FijarSubMun(r)
    FijarSupMun(r)
    FijarDefMunLo(r)
    FijarSubMunLo(r)
    FijarSupMunLo(r)
    LimiteDefRelMensual(r,m)
;

DefDeficit.. DefTotal =e= sum((r,m),Def(r,m));
DefGW..      GWTotal =e= sum((r,m),WSub(r,m));
DefSurface.. SurfaceTotal =e= sum((r,m),WSup(r,m));

BalanceAgua(r,m)..
    WSub(r,m)+WSup(r,m)+WAlm(r,m)+Def(r,m)
    =e= ReqActive(r,m)*Area(r);

LimiteSub(r)..
    sum(m,WSub(r,m)) =l= LimActivo(r,'LimSub_m3');
LimiteSup(r)..
    sum(m,WSup(r,m)) =l= LimActivo(r,'LimSup_m3');

* En E1 y los casos REPDA/H no hay topes mensuales de extraccion.
* A1 y A2 anaden la cota rho * LimAnual * diasMes/365.
LimiteSubMensual(r,m)$usaTopeMensual..
    WSub(r,m) =l= LimSubMes(r,m);
LimiteSupMensual(r,m)$usaTopeMensual..
    WSup(r,m) =l= LimSupMes(r,m);

LimiteCapt(r,m).. Capt(r,m) =l= CaptPotHa(r,m)*Area(r);

* Inventario ciclico: no se permite importar lluvia de otro municipio.
BalanceStoreEne(r)..
    Store(r,'Ene') =e= Store(r,'Dic')+Capt(r,'Ene')-WAlm(r,'Ene');
BalanceStore(r,m)$(ord(m)>1)..
    Store(r,m) =e= Store(r,m-1)+Capt(r,m)-WAlm(r,m);
LimiteStore(r,m)..
    Store(r,m) =l= CapAlmHa(r)*Area(r);

FijarDeficit.. DefTotal =l= DefStar+epsDef;
FijarSub..     GWTotal =l= SubStar+epsSub;

* La cuarta etapa conserva los totales regionales y municipales
* en bandas de +/-10 m3 para evitar problemas de escala numerica.
FijarDeficit4.. DefTotal =l= DefStar+epsDef4;
FijarSub4..     GWTotal =l= SubStar+epsSub4;
FijarSup4..     SurfaceTotal =l= SupStar+epsSup4;
FijarDeficitLo.. DefTotal =g= DefStar-epsDef4;
FijarSubLo..     GWTotal =g= SubStar-epsSub4;
FijarSupLo..     SurfaceTotal =g= SupStar-epsSup4;
FijarDefMun(r).. sum(m,Def(r,m)) =l= DefMunStar(r)+epsDef4;
FijarSubMun(r).. sum(m,WSub(r,m)) =l= SubMunStar(r)+epsSub4;
FijarSupMun(r).. sum(m,WSup(r,m)) =l= SupMunStar(r)+epsSup4;
FijarDefMunLo(r).. sum(m,Def(r,m)) =g= DefMunStar(r)-epsDef4;
FijarSubMunLo(r).. sum(m,WSub(r,m)) =g= SubMunStar(r)-epsSub4;
FijarSupMunLo(r).. sum(m,WSup(r,m)) =g= SupMunStar(r)-epsSup4;

LimiteDefRelMensual(r,m)$(ReqActive(r,m)>0)..
    Def(r,m) =l= MaxDefRel*ReqActive(r,m)*AreaMax(r);

Model MinDef /
    DefDeficit,DefGW,DefSurface,BalanceAgua,LimiteSub,LimiteSup,
    LimiteSubMensual,LimiteSupMensual,LimiteCapt,
    BalanceStoreEne,BalanceStore,LimiteStore /;
Model MinGW /
    DefDeficit,DefGW,DefSurface,BalanceAgua,LimiteSub,LimiteSup,
    LimiteSubMensual,LimiteSupMensual,LimiteCapt,
    BalanceStoreEne,BalanceStore,LimiteStore,FijarDeficit /;
Model MinSup /
    DefDeficit,DefGW,DefSurface,BalanceAgua,LimiteSub,LimiteSup,
    LimiteSubMensual,LimiteSupMensual,LimiteCapt,
    BalanceStoreEne,BalanceStore,LimiteStore,FijarDeficit,FijarSub /;
Model MinMaxDefRel /
    DefDeficit,DefGW,DefSurface,BalanceAgua,LimiteSub,LimiteSup,
    LimiteSubMensual,LimiteSupMensual,LimiteCapt,
    BalanceStoreEne,BalanceStore,LimiteStore,
    FijarDeficit4,FijarSub4,FijarSup4,
    FijarDeficitLo,FijarSubLo,FijarSupLo,
    FijarDefMun,FijarSubMun,FijarSupMun,
    FijarDefMunLo,FijarSubMunLo,FijarSupMunLo,
    LimiteDefRelMensual /;

* El escalamiento se utiliza solo en la cuarta etapa.
WSub.scale(r,m)=1e6;
WSup.scale(r,m)=1e6;
Capt.scale(r,m)=1e6;
WAlm.scale(r,m)=1e6;
Store.scale(r,m)=1e6;
Def.scale(r,m)=1e6;
Area.scale(r)=1e4;
DefTotal.scale=1e6;
GWTotal.scale=1e6;
SurfaceTotal.scale=1e6;
MaxDefRel.scale=1;
DefDeficit.scale=1e6;
DefGW.scale=1e6;
DefSurface.scale=1e6;
BalanceAgua.scale(r,m)=1e6;
LimiteSub.scale(r)=1e6;
LimiteSup.scale(r)=1e6;
LimiteSubMensual.scale(r,m)=1e6;
LimiteSupMensual.scale(r,m)=1e6;
LimiteCapt.scale(r,m)=1e6;
BalanceStoreEne.scale(r)=1e6;
BalanceStore.scale(r,m)=1e6;
LimiteStore.scale(r,m)=1e6;
FijarDeficit4.scale=1e6;
FijarSub4.scale=1e6;
FijarSup4.scale=1e6;
FijarDeficitLo.scale=1e6;
FijarSubLo.scale=1e6;
FijarSupLo.scale=1e6;
FijarDefMun.scale(r)=1e6;
FijarSubMun.scale(r)=1e6;
FijarSupMun.scale(r)=1e6;
FijarDefMunLo.scale(r)=1e6;
FijarSubMunLo.scale(r)=1e6;
FijarSupMunLo.scale(r)=1e6;
LimiteDefRelMensual.scale(r,m)=1e6;
MinMaxDefRel.scaleopt=1;

* Se mantienen las opciones numericas que se usaron al corregir E1.
$onecho > cplex.opt
eprhs 1e-8
epopt 1e-8
numericalemphasis 1
$offecho
MinMaxDefRel.optfile=1;

* Cada escenario se resuelve con las tres tecnologias.
* E1 siempre se ejecuta primero y se usa como referencia.
loop(s$sRun(s),

    fracCaptura=0.10;
    fracAlmacen=0.75;
    factorLim=1.00;
    rho=0;
    usaTopeMensual=0;

    if(sameas(s,'A1'),rho=2.0; usaTopeMensual=1;);
    if(sameas(s,'A2'),rho=1.5; usaTopeMensual=1;);
    if(sameas(s,'R80'),factorLim=0.80;);
    if(sameas(s,'R60'),factorLim=0.60;);
    if(sameas(s,'R40'),factorLim=0.40;);
    if(sameas(s,'R20'),factorLim=0.20;);
    if(sameas(s,'H0'),fracCaptura=0; fracAlmacen=0;);
    if(sameas(s,'H1'),fracCaptura=0.05;);
    if(sameas(s,'H2'),fracAlmacen=0.375;);
* H3 es adicional al manuscrito: sin almacen intermensual.
    if(sameas(s,'H3'),fracAlmacen=0;);

    LimActivo(r,lt)=factorLim*Lim(r,lt);
    LimSubMes(r,m)=0;
    LimSupMes(r,m)=0;
    if(usaTopeMensual=1,
        LimSubMes(r,m)=rho*Lim(r,'LimSub_m3')*dias(m)/365;
        LimSupMes(r,m)=rho*Lim(r,'LimSup_m3')*dias(m)/365;
    );

    ETcActiva(r,m)=ETcArchivo(r,m);
    if(sameas(s,'E2'),ETcActiva(r,m)=ETcE2(r,m););
    CaptPotHa(r,m)=10*max(0,Clima(r,m,'P_mm')-
                        Clima(r,m,'Pef_mm'))*fracCaptura;
    CapAlmHa(r)=fracAlmacen*sum(m,CaptPotHa(r,m));

    ParametrosEscenario(s,'FactorLimREPDA')=factorLim;
    ParametrosEscenario(s,'RhoMensual')=rho$usaTopeMensual;
    ParametrosEscenario(s,'Captura')=fracCaptura;
    ParametrosEscenario(s,'Almacen')=fracAlmacen;

    loop(kk,

        ReqActive(r,m)=10*max(0,ETcActiva(r,m)-
                             Clima(r,m,'Pef_mm'))/eta(kk);
        Area.fx(r)=AreaMax(r);
        Def.lo(r,m)=0;
        Def.up(r,m)=+inf;
        Def.up(r,m)$(ReqActive(r,m)<=0)=0;

        solve MinDef using lp minimizing DefTotal;
        abort$(MinDef.modelstat<>1 and MinDef.modelstat<>2)
            "Etapa 1 sin solucion optima",s,kk;
        DefStar=DefTotal.l;
        AuditoriaEtapas(s,kk,'Obj1Def_m3')=DefStar;
        AuditoriaEtapas(s,kk,'ModelStat1')=MinDef.modelstat;
        AuditoriaEtapas(s,kk,'SolveStat1')=MinDef.solvestat;

        solve MinGW using lp minimizing GWTotal;
        abort$(MinGW.modelstat<>1 and MinGW.modelstat<>2)
            "Etapa 2 sin solucion optima",s,kk;
        SubStar=GWTotal.l;
        AuditoriaEtapas(s,kk,'Obj2Sub_m3')=SubStar;
        AuditoriaEtapas(s,kk,'ModelStat2')=MinGW.modelstat;
        AuditoriaEtapas(s,kk,'SolveStat2')=MinGW.solvestat;

        solve MinSup using lp minimizing SurfaceTotal;
        abort$(MinSup.modelstat<>1 and MinSup.modelstat<>2)
            "Etapa 3 sin solucion optima",s,kk;
        SupStar=SurfaceTotal.l;
        AuditoriaEtapas(s,kk,'Obj3Sup_m3')=SupStar;
        AuditoriaEtapas(s,kk,'ModelStat3')=MinSup.modelstat;
        AuditoriaEtapas(s,kk,'SolveStat3')=MinSup.solvestat;

        DefMunStar(r)=sum(m,Def.l(r,m));
        SubMunStar(r)=sum(m,WSub.l(r,m));
        SupMunStar(r)=sum(m,WSup.l(r,m));

        solve MinMaxDefRel using lp minimizing MaxDefRel;
        abort$(MinMaxDefRel.modelstat<>1 and MinMaxDefRel.modelstat<>2)
            "Etapa 4 sin solucion optima",s,kk;
        AuditoriaEtapas(s,kk,'Obj4DefMax_pct')=100*MaxDefRel.l;
        AuditoriaEtapas(s,kk,'ModelStat4')=MinMaxDefRel.modelstat;
        AuditoriaEtapas(s,kk,'SolveStat4')=MinMaxDefRel.solvestat;

        MaxRelReal=smax((r,m)$(ReqActive(r,m)>0),
             Def.l(r,m)/(ReqActive(r,m)*AreaMax(r)));
        MaxBalance=smax((r,m),abs(WSub.l(r,m)+WSup.l(r,m)+
                        WAlm.l(r,m)+Def.l(r,m)-
                        ReqActive(r,m)*AreaMax(r)));
        MaxBalanceAlm=smax(r,abs(Store.l(r,'Ene')-Store.l(r,'Dic')-
                             Capt.l(r,'Ene')+WAlm.l(r,'Ene')));
        MaxDefMun=smax(r,abs(sum(m,Def.l(r,m))-DefMunStar(r)));
        MaxSubMun=smax(r,abs(sum(m,WSub.l(r,m))-SubMunStar(r)));
        MaxSupMun=smax(r,abs(sum(m,WSup.l(r,m))-SupMunStar(r)));
        MaxDefReg=abs(DefTotal.l-DefStar);
        MaxSubReg=abs(GWTotal.l-SubStar);
        MaxSupReg=abs(SurfaceTotal.l-SupStar);

        AuditoriaEtapas(s,kk,'BalanceMax_m3')=MaxBalance;
        AuditoriaEtapas(s,kk,'AlmacenMax_m3')=MaxBalanceAlm;
        AuditoriaEtapas(s,kk,'DeltaDefMunicipal_m3')=MaxDefMun;
        AuditoriaEtapas(s,kk,'DeltaSubMunicipal_m3')=MaxSubMun;
        AuditoriaEtapas(s,kk,'DeltaSupMunicipal_m3')=MaxSupMun;
        AuditoriaEtapas(s,kk,'DeltaDefRegional_m3')=MaxDefReg;
        AuditoriaEtapas(s,kk,'DeltaSubRegional_m3')=MaxSubReg;
        AuditoriaEtapas(s,kk,'DeltaSupRegional_m3')=MaxSupReg;
        AuditoriaEtapas(s,kk,'DiferenciaObjetivo4_pct')=
            100*abs(MaxDefRel.l-MaxRelReal);

        abort$(MaxBalance>0.1 or MaxBalanceAlm>0.1)
            "No cierra el balance de agua",s,kk;
        abort$(MaxDefMun>epsDef4+1 or MaxSubMun>epsSub4+1 or
               MaxSupMun>epsSup4+1)
            "La cuarta etapa modifica las cantidades municipales",s,kk;
        abort$(MaxDefReg>epsDef4+1 or MaxSubReg>epsSub4+1 or
               MaxSupReg>epsSup4+1)
            "La cuarta etapa modifica las cantidades regionales",s,kk;
        abort$(abs(MaxDefRel.l-MaxRelReal)>1e-5)
            "No coincide la funcion objetivo de la cuarta etapa",s,kk;

* ETa estimada se obtiene convirtiendo el deficit de riego
* bruto a un deficit neto, con la eficiencia de cada tecnologia.
        ETmAnual(r)=10*AreaMax(r)*sum(m,ETcActiva(r,m));
        ETaEstimada(r)=ETmAnual(r)-eta(kk)*sum(m,Def.l(r,m));
        abort$(smin(r,ETmAnual(r))<=0 or
               smin(r,ETaEstimada(r))< -0.001)
            "Evapotranspiracion estimada fuera de rango",s,kk;
        EstresET(r)=max(0,min(1,
            (ETmAnual(r)-ETaEstimada(r))/ETmAnual(r)));

        GestionMensual(r,m,s,kk,'Demanda_m3')=
            ReqActive(r,m)*AreaMax(r);
        GestionMensual(r,m,s,kk,'Subterranea_m3')=WSub.l(r,m);
        GestionMensual(r,m,s,kk,'Superficial_m3')=WSup.l(r,m);
        GestionMensual(r,m,s,kk,'LluviaUsada_m3')=WAlm.l(r,m);
        GestionMensual(r,m,s,kk,'Deficit_m3')=Def.l(r,m);
        GestionMensual(r,m,s,kk,'Almacen_m3')=Store.l(r,m);
        GestionMensual(r,m,s,kk,'Cobertura_pct')$(ReqActive(r,m)>0)=
            100*(1-Def.l(r,m)/(ReqActive(r,m)*AreaMax(r)));
        GestionMensual(r,m,s,kk,'ETm_m3')=
            10*AreaMax(r)*ETcActiva(r,m);
        GestionMensual(r,m,s,kk,'ETaEstimada_m3')=
            10*AreaMax(r)*ETcActiva(r,m)-eta(kk)*Def.l(r,m);

        GestionMunicipal(r,s,kk,'Area_ha')=AreaMax(r);
        GestionMunicipal(r,s,kk,'ProduccionReferencia_ton')=
            Hist(r,'Prod_ton');
        GestionMunicipal(r,s,kk,'Demanda_m3')=
            sum(m,ReqActive(r,m)*AreaMax(r));
        GestionMunicipal(r,s,kk,'Subterranea_m3')=
            sum(m,WSub.l(r,m));
        GestionMunicipal(r,s,kk,'Superficial_m3')=
            sum(m,WSup.l(r,m));
        GestionMunicipal(r,s,kk,'LluviaCaptada_m3')=
            sum(m,Capt.l(r,m));
        GestionMunicipal(r,s,kk,'LluviaUsada_m3')=
            sum(m,WAlm.l(r,m));
        GestionMunicipal(r,s,kk,'Deficit_m3')=
            sum(m,Def.l(r,m));
        GestionMunicipal(r,s,kk,'Cobertura_pct')
            $(GestionMunicipal(r,s,kk,'Demanda_m3')>0)=
            100*(1-GestionMunicipal(r,s,kk,'Deficit_m3')/
                   GestionMunicipal(r,s,kk,'Demanda_m3'));
        GestionMunicipal(r,s,kk,'PresionSub_pct')
            $(LimActivo(r,'LimSub_m3')>0)=
            100*GestionMunicipal(r,s,kk,'Subterranea_m3')/
                LimActivo(r,'LimSub_m3');
        GestionMunicipal(r,s,kk,'PresionSup_pct')
            $(LimActivo(r,'LimSup_m3')>0)=
            100*GestionMunicipal(r,s,kk,'Superficial_m3')/
                LimActivo(r,'LimSup_m3');
        GestionMunicipal(r,s,kk,'MaxDefRelMensual_pct')=
            smax(m$(ReqActive(r,m)>0),
                 100*Def.l(r,m)/(ReqActive(r,m)*AreaMax(r)));
        GestionMunicipal(r,s,kk,'ETm_m3')=ETmAnual(r);
        GestionMunicipal(r,s,kk,'ETaEstimada_m3')=ETaEstimada(r);
        GestionMunicipal(r,s,kk,'EstresET_pct')=100*EstresET(r);

        GestionRegional(s,kk,'Area_ha')=sum(r,AreaMax(r));
        GestionRegional(s,kk,'ProduccionReferencia_ton')=
            sum(r,Hist(r,'Prod_ton'));
        GestionRegional(s,kk,'Demanda_m3')=
            sum((r,m),ReqActive(r,m)*AreaMax(r));
        GestionRegional(s,kk,'Subterranea_m3')=
            sum((r,m),WSub.l(r,m));
        GestionRegional(s,kk,'Superficial_m3')=
            sum((r,m),WSup.l(r,m));
        GestionRegional(s,kk,'LluviaCaptada_m3')=
            sum((r,m),Capt.l(r,m));
        GestionRegional(s,kk,'LluviaUsada_m3')=
            sum((r,m),WAlm.l(r,m));
        GestionRegional(s,kk,'Deficit_m3')=
            sum((r,m),Def.l(r,m));
        GestionRegional(s,kk,'Cobertura_pct')
            $(GestionRegional(s,kk,'Demanda_m3')>0)=
            100*(1-GestionRegional(s,kk,'Deficit_m3')/
                   GestionRegional(s,kk,'Demanda_m3'));
        GestionRegional(s,kk,'MaxDefRelMensual_pct')=
            100*MaxDefRel.l;
        GestionRegional(s,kk,'ETm_m3')=sum(r,ETmAnual(r));
        GestionRegional(s,kk,'ETaEstimada_m3')=sum(r,ETaEstimada(r));
        GestionRegional(s,kk,'EstresET_pct')=
            100*(1-sum(r,ETaEstimada(r))/sum(r,ETmAnual(r)));

        loop(y,
            ProduccionKyMunicipal(r,s,kk,y,'Ky')=kyValue(y);
            ProduccionKyMunicipal(r,s,kk,y,'EstresET_pct')=
                100*EstresET(r);
            ProduccionKyMunicipal(r,s,kk,y,'CoberturaRiego_pct')=
                GestionMunicipal(r,s,kk,'Cobertura_pct');
            ProduccionKyMunicipal(r,s,kk,y,'ProduccionReferencia_ton')=
                Hist(r,'Prod_ton');
            ProduccionKyMunicipal(r,s,kk,y,'ProduccionAjustada_ton')=
                Hist(r,'Prod_ton')*max(0,1-kyValue(y)*EstresET(r));
            ProduccionKyMunicipal(r,s,kk,y,'PerdidaProduccion_ton')=
                Hist(r,'Prod_ton')-
                ProduccionKyMunicipal(r,s,kk,y,'ProduccionAjustada_ton');
            ProduccionKyMunicipal(r,s,kk,y,'FactorRendimiento_pct')=
                100*ProduccionKyMunicipal(r,s,kk,y,'ProduccionAjustada_ton')/
                    Hist(r,'Prod_ton');

            ProduccionKyRegional(s,kk,y,'Ky')=kyValue(y);
            ProduccionKyRegional(s,kk,y,'ProduccionReferencia_ton')=
                sum(r,Hist(r,'Prod_ton'));
            ProduccionKyRegional(s,kk,y,'ProduccionAjustada_ton')=
                sum(r,ProduccionKyMunicipal(r,s,kk,y,
                                           'ProduccionAjustada_ton'));
            ProduccionKyRegional(s,kk,y,'PerdidaProduccion_ton')=
                sum(r,ProduccionKyMunicipal(r,s,kk,y,
                                           'PerdidaProduccion_ton'));
            ProduccionKyRegional(s,kk,y,'FactorRendimiento_pct')=
                100*ProduccionKyRegional(s,kk,y,'ProduccionAjustada_ton')/
                    ProduccionKyRegional(s,kk,y,'ProduccionReferencia_ton');
            ProduccionKyRegional(s,kk,y,'EstresET_pct')=
                GestionRegional(s,kk,'EstresET_pct');
            ProduccionKyRegional(s,kk,y,'CoberturaRiego_pct')=
                GestionRegional(s,kk,'Cobertura_pct');
        );
    );
);

* Se compara cada corrida con E1 para la misma tecnologia.
loop(s$sRun(s),
    ResumenComparativo(r,s,k,'Demanda_m3')=
        GestionMunicipal(r,s,k,'Demanda_m3');
    ResumenComparativo(r,s,k,'Deficit_m3')=
        GestionMunicipal(r,s,k,'Deficit_m3');
    ResumenComparativo(r,s,k,'Cobertura_pct')=
        GestionMunicipal(r,s,k,'Cobertura_pct');
    ResumenComparativo(r,s,k,'DeltaDeficit_m3')=
        GestionMunicipal(r,s,k,'Deficit_m3')-
        GestionMunicipal(r,'E1',k,'Deficit_m3');
    ResumenComparativo(r,s,k,'DeltaCobertura_pp')=
        GestionMunicipal(r,s,k,'Cobertura_pct')-
        GestionMunicipal(r,'E1',k,'Cobertura_pct');
    ResumenComparativo(r,s,k,'CoberturaCompleta')=
        1$(GestionMunicipal(r,s,k,'Deficit_m3')<=epsDef4+1);
    ResumenComparativo(r,s,k,'DeficitRelMax_pct')=
        GestionMunicipal(r,s,k,'MaxDefRelMensual_pct');
    ResumenComparativo(r,s,k,'EstresET_pct')=
        GestionMunicipal(r,s,k,'EstresET_pct');
);

* Los archivos GDX conservan los resultados con toda su precision.
execute_unload 'RESULTADOS_WF_NEXUS_E1_INTEGRADO.gdx',
    GestionRegional, GestionMunicipal, GestionMensual,
    ProduccionKyRegional, ProduccionKyMunicipal,
    AuditoriaEtapas, ResumenComparativo,
    ParametrosEscenario, EscenariosEjecutados, kyValue, eta, EToMensual, KcEfectivo,
    ETcArchivo, ETcE2, Clima, AreaMax, Hist, Lim;

$onecho > salida_wf_e1_ky.txt
par=GestionRegional       rng=GestionRegional!A1       rdim=2 cdim=1
par=GestionMunicipal      rng=GestionMunicipal!A1      rdim=3 cdim=1
par=GestionMensual        rng=GestionMensual!A1        rdim=4 cdim=1
par=ProduccionKyRegional  rng=ProduccionKyRegional!A1  rdim=3 cdim=1
par=ProduccionKyMunicipal rng=ProduccionKyMunicipal!A1 rdim=4 cdim=1
par=AuditoriaEtapas       rng=AuditoriaEtapas!A1       rdim=2 cdim=1
par=ResumenComparativo    rng=ResumenComparativo!A1    rdim=3 cdim=1
par=ParametrosEscenario   rng=ParametrosEscenario!A1   rdim=1 cdim=1
par=EscenariosEjecutados  rng=EscenariosEjecutados!A1  rdim=1 cdim=0
par=kyValue              rng=ValoresKy!A1            rdim=1 cdim=0
par=eta                  rng=Eficiencias!A1          rdim=1 cdim=0
$offecho
execute.checkErrorLevel 'gdxxrw RESULTADOS_WF_NEXUS_E1_INTEGRADO.gdx output=%OUTPUT% @salida_wf_e1_ky.txt trace=3';

Display GIR70Calculado, GIR70Esperado, GIR70E2Calculado,
        AuditoriaEtapas, GestionRegional, ProduccionKyRegional;
