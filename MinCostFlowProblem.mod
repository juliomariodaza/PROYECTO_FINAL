# ------------------------------------------------------------------------------------------- #
# -------------------------------- Minimum Cost Flow Problem -------------------------------- #
# ------------------------------------------------------------------------------------------- #
# -------------------------------- Julio Mario Daza Escorcia -------------------------------- #
# ------------------------------------------------------------------------------------------- #

reset;

param NODOS;								# Número de Nodos del Grafo.
set ARCOS within {1..NODOS, 1..NODOS};		# ARCOS contenidos (within) entre los Nodos.
param COSTO{ARCOS};							# Costo de los Arcos (i,j).
param CAPACIDAD{ARCOS};						# Capacidad de los Arcos (i, j).
param DEMANDA{1..NODOS};					# Demanda de los Nodos.
check: sum{i in 1..NODOS} DEMANDA[i] = 0;	# Verificar que la suma de las demandas sea cero.

var X{ARCOS} >=0;							# Definición de Variables de Flujo X_ij.

minimize COSTO_TOTAL:						# Definición de Función Objetivo (FO): Min MCF.
      		sum{(i,j) in ARCOS} COSTO[i,j] * X[i,j];

subject to BALANCE {i in 1..NODOS}:			# Restricción de Balance (Int Flow == Out Flow).
  			sum {(i,j) in ARCOS} X[i,j] - sum{(j,i) in ARCOS} X[j,i] = DEMANDA[i];

subject to CAPACIDAD_ARCOS {(i,j) in ARCOS}:# Restricción de Capacidad.
 			X[i,j] <= CAPACIDAD[i,j];		# Limita el Flujo Máximo en los Arcos (i,j).

# ------------------------------------------------------------------------------------------- #
# -------------------------------- Minimum Cost Flow Problem -------------------------------- #
# ------------------------------------------------------------------------------------------- #
# -------------------------------- Julio Mario Daza Escorcia -------------------------------- #
# ------------------------------------------------------------------------------------------- #