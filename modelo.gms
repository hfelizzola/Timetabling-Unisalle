$title   Timetabling

sets
i        Franjas Horarias /F1*F11/
j        Materias /M1*M90/
k        Profesores disponibles/P1*P30/
n        Dias habiles/D1*D6/
m        Semestre al que puede pertenecer una materia/S1*S10/
t        Numero de pares de franjas horarias de dos horas /T1*T5/
q        Numero de trios de franjas horarias de tres horas /Q1*Q3/;

Parameter CFM(n,i,J);
*Matriz de 1 y O si en la franja i en el dia n esta disponible el profesor k
$CALL GDXXRW parametros.xlsx par=CFM rng=CFM!A1:CN67 Rdim=2 Cdim=1
$GDXIN parametros.gdx
$LOAD CFM
$GDXIN

Parameter MP(j,k);
*Matriz de 1 y 0 1 si la materia j pertenece al semestre m 0 sino
$CALL GDXXRW parametros.xlsx par=MP rng=MP!A1:AE91 Cdim=1 Rdim=1
$GDXIN parametros.gdx
$LOAD MP
$gdxIn

parameter FP(n,i,k);
*Matriz de 1 y O si en la franja i en el dia n esta disponible el profesor k
$CALL GDXXRW parametros.xlsx par=FP rng=FP!A1:AF67 Rdim=2 Cdim=1
$GDXIN parametros.gdx
$LOAD FP
$GDXIN

Parameter MS(j,m);
*Matriz de 1 y 0 1 si la materia j pertenece al semestre m 0 sino
$CALL GDXXRW parametros.xlsx par=MS rng=MS!A1:L91 Cdim=1 Rdim=1
$GDXIN parametros.gdx
$LOAD MS
$gdxIn

Parameter INTH(j)
*Numero de horas que debe tener la materia j
$CALL GDXXRW parametros.xlsx par=INTH rng=INTH!A1:B90 Cdim=0 Rdim=1
$GDXIN parametros.gdx
$LOAD INTH
$gdxIn

Parameter PMAX(K)
*Numero de horas que debe tener la materia j
$CALL GDXXRW parametros.xlsx par=PMAX rng=PMAX!A1:B30 Cdim=0 Rdim=1
$GDXIN parametros.gdx
$LOAD PMAX
$gdxIn

Parameter PMIN(K)
*Numero de horas que debe tener la materia j
$CALL GDXXRW parametros.xlsx par=PMIN rng=PMIN!A1:B30 Cdim=0 Rdim=1
$GDXIN parametros.gdx
$LOAD PMIN
$gdxIn


Parameter DOSF(J)
*Numero de horas que debe tener la materia j
$CALL GDXXRW parametros.xlsx par=DOSF rng=DOSF!A1:B90 Cdim=0 Rdim=1
$GDXIN parametros.gdx
$LOAD DOSF
$gdxIn

Parameter TRESF(J)
*Numero de horas que debe tener la materia j
$CALL GDXXRW parametros.xlsx par=TRESF rng=TRESF!A1:B90 Cdim=0 Rdim=1
$GDXIN parametros.gdx
$LOAD TRESF
$gdxIn

Binary Variable
X(i,j,k,n)       1 si en la franja horaria i asigno la materia j al profesor k en el dia k 0 si no
Y(j,k)           1 si el profe k va a dictar la materia j
H(j,k,n,t)       1 si se activa el grupo de franjas para la materia 0 si no
G(j,k,n,q)        1 si se activa la materia j que dicta el profesor k, el dia n en el tramo de tres horas q;

Free Variable
F        suma de todas las posibles asignaciones multiplicada por la matriz de costos de las franjas horaria;

Equations
FO                       Funcion Objetivo
RDOC(i,j,n)              Garantiza que solo se asigne maximo un docente a una materia en una franja y dia respectivos. Restriccion dura
RASIG(i,k,n)             Garantiza que una asignatura solo se pueda programar maximo una vez en una franja horaria. Restriccion dura
RINTH(j)                 Toda asignatura debe ser programada seg?n su intensidad horaria semanal. Restriccion dura (Ecuación 4)
RCAPS(j,k)               Garantiza que cada asignatura sea dictada por un unico docente. Restriccion dura (Ecuación 5)
DISP (N,I,K,J)           Esta ecuación verifica que el docente tiene la disponibilidad de tiempo para dictar la asignatura (ecuación 6)
SABE (i,j,k,n)           Garantiza que el profesor maneja la asignatura asignada (ecuacion 7)
*MAXHOR (K,N)             MAXIMO DE HORAS QUE PUEDE DICTAR UN DOCENTE POR DIA (Ecuación 8)
SEMESTRE (i,m,n)         Garantiza que si hay 2 materias del mismo semestre. no se hagan a la misma hora (Ecuación 9) RESTRICCION CON PROBLEMAS
RHSMAX(k)                Garantiza que un docente no dicte mas de 18 horas semanales (Ecuación 10)
RHSMIN(k)                Garantiza que un docente no dicte MENOS DE TANTAS horas semanales (Ecuación 11)
RFUDOSH(i,j,k,n)         Restriccion para la franja 1 de dos horas (Ecuación12)
RFDDOSH(i,j,k,n)         Restriccion para la franja 2 de dos horas (Ecuación 13)
RFTDOSH(i,j,k,n)         Restriccion para franja 3 de 2 horas (Ecuación 14)
RFCDOSH(i,j,k,n)         Restriccion para franja 4 de 2 horas (Ecuación 15)
RFCIDOSH(i,j,k,n)        Restriccion para franja 5 de 2 horas (Ecuación 16)
ALMUERZO (i,j,k,n)       Restricción que garantiza una franja de una hora para el almuerzo (Ecuación 17)

RDOC2(j,n)              Garantiza que una asignatura solo se de en una franja horaria por diA
RFUTRESH(i,j,k,n)       Restriccion para la franja 1 de tres horas tiene problemas vuelve todo cero (Ecuación 19)
RFDTRESH(i,j,k,n)       Restriccion para la franja 2 de tres horas (Ecuación 20)
RFTTRESH(i,j,k,n)       Restriccion para la franja 3 de tres horas (Ecuación 21)
TRESU(j)                Materia de tres horas el mismo dia
CAP(i,n)                No puede haber mas de 9 salones al tiempo
;

FO..                             F=e=sum((i,j,k,n),CFM(n,i,j)*X(i,j,k,n));
RDOC(i,j,n)..                    sum(k,X(i,j,k,n))=l=1;
RASIG(i,k,n)..                   sum(j,X(i,j,k,n))=l=1;
RINTH(j)..                       sum((i,k,n),X(i,j,k,n))=e=INTH(j);
RCAPS(j,k)..                     sum((i,n),X(i,j,k,n))=e=INTH(j)*Y(j,k);
RHSMAX(k)..                      sum((i,j,n),X(i,j,k,n))=l=PMAX(k);
RHSMIN(k)..                      sum((i,j,n),X(i,j,k,n))=g=PMIN(k);
SABE(i,j,k,n)..                  X(i,j,k,n)=l=MP(j,k);
SEMESTRE (i,m,n)..               sum((j,k)$(MS(j,m)=1), X(i,j,k,n)*MP(j,k))=l=1;
DISP (N,I,K,J)..                 X(i,j,k,n)=l=FP(n,i,k);
RFUDOSH(i,j,k,n)$(DOSF(j)=1)..   X('F1',j,k,n)+X('F2',j,k,n)=e=2*H(j,k,n,'T1');
RFDDOSH(i,j,k,n)$(DOSF(j)=1)..   X('F3',j,k,n)+X('F4',j,k,n)=e=2*H(j,k,n,'T2');
RFTDOSH(i,j,k,n)$(DOSF(j)=1)..   X('F5',j,k,n)+X('F6',j,k,n)=e=2*H(j,k,n,'T3');
RFCDOSH(i,j,k,n)$(DOSF(j)=1)..   X('F8',j,k,n)+X('F9',j,k,n)=e=2*H(j,k,n,'T4');
RFCIDOSH(i,j,k,n)$(DOSF(j)=1)..  X('F10',j,k,n)+X('F11',j,k,n)=e=2*H(j,k,n,'T5');
ALMUERZO(i,j,k,n)..              X('F7',j,k,n)=e=0;
RDOC2(j,n)$(DOSF(j)=1)..         sum((k,t),H(j,k,n,t))=l=1;
RFUTRESH(i,j,k,n)$(DOSF(j)=0)..  X('F1',j,k,n)+X('F2',j,k,n)+ X('F3',j,k,n)=e=3*G(j,k,n,'Q1');
RFDTRESH(i,j,k,n)$(DOSF(j)=0)..  X('F4',j,k,n)+ X('F5',j,k,n)+X('F6',j,k,n)=e=3*G(j,k,n,'Q2');
RFTTRESH(i,j,k,n)$(DOSF(j)=0)..  X('F8',j,k,n)+X('F9',j,k,n)+X('F10',j,k,n)=e=3*G(j,k,n,'Q3');
TRESU(j)$(DOSF(j)=0)..           sum((k,n,q),G(j,k,n,q))=e=1;
CAP(i,n)..                       sum((j,k),X(i,j,k,n))=l=9;

model Timetabling /all/;
option MIP=CBC;
solve Timetabling using MIP minimizing F;
display X.l;
display PMIN;
display Y.l;
display F.l;
display H.l;
display G.l;

execute_unload "resultados.gdx" X.l y.l f.l h.l;
execute 'gdxxrw.exe resultados.gdx o=resultados.xlsx var=x.L'