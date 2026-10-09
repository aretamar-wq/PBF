#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Agrega cuentas nuevas (pesos y dólares) a un PBF de BASE24 y genera un archivo
válido para el refresh (registros en orden estricto por clave, posiciones 12-36).

Solo usa la librería estándar.

Ejemplo de ejecución:
    python3 pbf_agregar_cuentas.py PBF_ENTRADA.dat
    python3 pbf_agregar_cuentas.py PBF_ENTRADA.dat --salida PBF_NUEVO.dat \
        --cantidad 100000 --fecha-header 261009 --fecha-detalle 261008
    python3 pbf_agregar_cuentas.py PBF_ENTRADA.dat --cuenta-inicial 80000 \
        --umbral-serie 1000000
"""
import argparse
import re
import sys
from collections import Counter
from datetime import date, timedelta

LONG = 116

# Posiciones 0-based (slices de Python) derivadas de las 1-based del enunciado
CLAVE = slice(11, 36)        # 12-36 clave de ordenamiento
GRUPO = slice(15, 22)        # 16-22 "00105" + moneda -> grupo 0010500 / 0010502
CUENTA = slice(22, 33)       # 23-33 número de cuenta
SALDO2 = slice(50, 62)       # 51-62 saldo 2

GRUPO_PESOS = "0010500"
GRUPO_DOLARES = "0010502"

# Overpunch: { A-I = positivo 0-9 ; } J-R = negativo 0-9
_POS = {"{": 0, **{c: i + 1 for i, c in enumerate("ABCDEFGHI")}}
_NEG = {"}": 0, **{c: i + 1 for i, c in enumerate("JKLMNOPQR")}}
_POS_INV = {v: k for k, v in _POS.items()}
_NEG_INV = {v: k for k, v in _NEG.items()}


def decodificar_zoned(campo):
    """Convierte un campo zoned decimal (overpunch) a entero con signo."""
    ultimo = campo[-1]
    cuerpo = campo[:-1]
    if ultimo in _POS:
        signo, d = 1, _POS[ultimo]
    elif ultimo in _NEG:
        signo, d = -1, _NEG[ultimo]
    elif ultimo.isdigit():          # numérico plano (sin overpunch)
        signo, d = 1, int(ultimo)
    else:
        raise ValueError("campo zoned inválido: %r" % campo)
    if cuerpo and not cuerpo.isdigit():
        raise ValueError("campo zoned inválido: %r" % campo)
    return signo * int((cuerpo or "0") + str(d))


def codificar_zoned(valor, ancho):
    """Inverso de decodificar_zoned: entero -> campo overpunch de 'ancho' chars."""
    s = str(abs(valor)).zfill(ancho)
    if len(s) > ancho:
        raise ValueError("el valor %d no entra en %d posiciones" % (valor, ancho))
    tabla = _NEG_INV if valor < 0 else _POS_INV
    return s[:-1] + tabla[int(s[-1])]


def fecha_aammdd(texto):
    if not re.fullmatch(r"\d{6}", texto):
        raise argparse.ArgumentTypeError("la fecha debe ser AAMMDD (6 dígitos)")
    return texto


def error(msg):
    sys.exit("ERROR: " + msg)


def leer_registros(ruta):
    with open(ruta, "rb") as f:
        datos = f.read()
    if len(datos) == 0 or len(datos) % LONG != 0:
        error("el tamaño del archivo (%d) no es múltiplo de %d" % (len(datos), LONG))
    if b"\n" in datos or b"\r" in datos:
        error("el archivo contiene saltos de línea; se esperaba todo concatenado")
    texto = datos.decode("ascii")
    return [texto[i:i + LONG] for i in range(0, len(texto), LONG)]


def validar_entrada(regs):
    """Valida estructura FH, BH, detalles, BT, FT y contadores. Devuelve detalles."""
    if len(regs) < 4:
        error("el archivo tiene menos de 4 registros")
    if regs[0][7:9] != "FH":
        error("el primer registro no es FH")
    if regs[1][7:9] != "BH":
        error("el segundo registro no es BH")
    if regs[-2][7:9] != "BT":
        error("el anteúltimo registro no es BT")
    if regs[-1][7:9] != "FT":
        error("el último registro no es FT")
    detalles = regs[2:-2]
    for i, r in enumerate(detalles, start=3):
        if not r.startswith("0116"):
            error("el registro %d no es de detalle (no empieza con 0116)" % i)
    bt, ft = regs[-2], regs[-1]
    if int(bt[25:32]) != len(detalles):
        error("BT declara %s detalles y hay %d" % (bt[25:32], len(detalles)))
    if int(ft[9:16]) != len(detalles):
        error("FT declara %s detalles y hay %d" % (ft[9:16], len(detalles)))
    total = sum(decodificar_zoned(r[SALDO2]) for r in detalles)
    try:
        total_bt = decodificar_zoned(bt[9:25])
        if total_bt != total:
            print("ADVERTENCIA: total BT de entrada (%d) != suma de saldos (%d)"
                  % (total_bt, total), file=sys.stderr)
    except ValueError:
        print("ADVERTENCIA: total BT de entrada ilegible: %r" % bt[9:25],
              file=sys.stderr)
    return detalles


def armar_detalle(moneda, tipo, cuenta, fecha_detalle):
    """Arma un registro nuevo de detalle (seq provisoria 0000000, se renumera luego)."""
    r = ("0116" + "0000000" + "0312" + "00105" + moneda + str(cuenta).zfill(11)
         + " " + tipo + "IF"
         + "00000000000{00000000000{000000000000000000000099000000000000"
         + fecha_detalle + "0000000000  ")
    assert len(r) == LONG, "registro generado de largo %d (esperado %d)" % (len(r), LONG)
    return r


def main():
    ayer = (date.today() - timedelta(days=1)).strftime("%y%m%d")
    hoy = date.today().strftime("%y%m%d")

    ap = argparse.ArgumentParser(description="Agrega cuentas nuevas a un PBF BASE24.")
    ap.add_argument("archivo", help="archivo PBF de entrada")
    ap.add_argument("--salida", help="archivo de salida (default: entrada + '_out')")
    ap.add_argument("--cantidad", type=int, default=100000,
                    help="cuentas a agregar por moneda (default 100000)")
    ap.add_argument("--cuenta-inicial", type=int,
                    help="forzar la cuenta de arranque")
    ap.add_argument("--umbral-serie", type=int, default=1000000,
                    help="cuentas >= umbral se consideran fuera de serie (default 1000000)")
    ap.add_argument("--fecha-header", type=fecha_aammdd, default=hoy,
                    help="fecha del FH, AAMMDD (default hoy)")
    ap.add_argument("--fecha-detalle", type=fecha_aammdd, default=ayer,
                    help="fecha de los detalles nuevos, AAMMDD (default ayer)")
    args = ap.parse_args()
    if args.cantidad < 1:
        error("--cantidad debe ser >= 1")
    salida = args.salida or args.archivo + "_out"

    # 1. Lectura y validación
    regs = leer_registros(args.archivo)
    detalles = validar_entrada(regs)
    fh, bh = regs[0], regs[1]

    # 2. Cuenta inicial (ignorando cuentas sueltas fuera de serie)
    max_p = max_d = None
    fuera_de_serie = []
    for r in detalles:
        g = r[GRUPO]
        if g != GRUPO_PESOS and g != GRUPO_DOLARES:
            continue
        cta = int(r[CUENTA])
        if cta >= args.umbral_serie:
            fuera_de_serie.append((g, cta))
            continue
        if g == GRUPO_PESOS:
            if max_p is None or cta > max_p:
                max_p = cta
        elif max_d is None or cta > max_d:
            max_d = cta
    maximos = [m for m in (max_p, max_d) if m is not None]
    if args.cuenta_inicial is not None:
        inicio = args.cuenta_inicial
    elif maximos:
        inicio = max(maximos) + 1
    else:
        error("no hay cuentas en los grupos 0010500/0010502; use --cuenta-inicial")
    fin = inicio + args.cantidad - 1
    if inicio < 0 or fin > 99999999999:
        error("rango de cuentas fuera de 11 dígitos: %d-%d" % (inicio, fin))

    # 3. Generación de las cuentas nuevas
    nuevos_p = [armar_detalle("00", "11", c, args.fecha_detalle)
                for c in range(inicio, fin + 1)]
    nuevos_d = [armar_detalle("02", "15", c, args.fecha_detalle)
                for c in range(inicio, fin + 1)]
    nuevos = nuevos_p + nuevos_d

    # 4. Duplicados: contra el archivo y entre sí (set => O(1) por consulta)
    claves_exist = {r[CLAVE] for r in detalles}
    claves_nuevas = {r[CLAVE] for r in nuevos}
    dup = claves_exist & claves_nuevas
    if dup or len(claves_nuevas) != len(nuevos):
        ej = sorted(dup)[:5]
        error("claves duplicadas (%d ya existen en el archivo). Ejemplos: %s"
              % (len(dup), ej))

    # 5. Unir y ordenar por clave con un solo sort
    todos = detalles + nuevos
    todos.sort(key=lambda r: r[CLAVE])

    # 6. Renumerar secuencia: FH=1, BH=2, detalles desde 3
    pos_nuevas = {"00": [None, None], "02": [None, None]}   # [primer, ultimo] nº registro
    total_saldos = 0
    out_det = []
    for i, r in enumerate(todos):
        nro = i + 3
        if r[CLAVE] in claves_nuevas:
            rango = pos_nuevas[r[20:22]]
            if rango[0] is None:
                rango[0] = nro
            rango[1] = nro
        total_saldos += decodificar_zoned(r[SALDO2])
        out_det.append(r[:4] + "%07d" % nro + r[11:])

    n_det = len(out_det)

    # 7-8. Header, BH, trailers
    out_fh = "%07d" % 1 + fh[7:16] + args.fecha_header + fh[22:]
    out_bh = "%07d" % 2 + bh[7:]
    bt = regs[-2]
    ft = regs[-1]
    # El total del BT conserva el formato de la entrada: dígitos planos si el
    # original los usaba (y el total no es negativo); si no, overpunch.
    if bt[24].isdigit() and total_saldos >= 0:
        total_txt = str(total_saldos).zfill(16)
    else:
        total_txt = codificar_zoned(total_saldos, 16)
    out_bt = ("%07d" % (n_det + 3) + bt[7:9] + total_txt
              + "%07d" % n_det + bt[32:])
    out_ft = "%07d" % (n_det + 4) + ft[7:9] + "%07d" % n_det + ft[16:]
    salida_regs = [out_fh, out_bh] + out_det + [out_bt, out_ft]

    # Validaciones finales
    assert all(len(r) == LONG for r in salida_regs), "hay registros que no miden 116"
    claves = [r[CLAVE] for r in out_det]
    for a, b in zip(claves, claves[1:]):
        if not a < b:
            error("claves fuera de orden o repetidas: %r >= %r" % (a, b))
    assert int(out_bt[25:32]) == n_det and int(out_ft[9:16]) == n_det
    assert decodificar_zoned(out_bt[9:25]) == sum(
        decodificar_zoned(r[SALDO2]) for r in out_det)
    # Originales intactos (ignorando secuencia): comparar como multiconjuntos
    orig = Counter(r[:4] + r[11:] for r in detalles)
    conservados = Counter(r[:4] + r[11:] for r in out_det
                          if r[CLAVE] not in claves_nuevas)
    if orig != conservados:
        error("los registros originales no se conservaron intactos")
    cuerpo = "".join(salida_regs)
    assert len(cuerpo) == LONG * len(salida_regs) and "\n" not in cuerpo

    with open(salida, "wb") as f:
        f.write(cuerpo.encode("ascii"))

    # Resumen por consola
    fmt = lambda v: "ninguna" if v is None else str(v)
    print("Archivo de entrada : %s (%d registros, %d detalles)"
          % (args.archivo, len(regs), len(detalles)))
    print("Cuenta más alta pesos   (0010500): %s" % fmt(max_p))
    print("Cuenta más alta dólares (0010502): %s" % fmt(max_d))
    if fuera_de_serie:
        print("ADVERTENCIA - cuentas fuera de serie ignoradas (>= %d), dejadas sin modificar:"
              % args.umbral_serie)
        for g, c in fuera_de_serie:
            print("    grupo %s cuenta %d" % (g, c))
    else:
        print("Cuentas fuera de serie ignoradas: ninguna")
    print("Rango generado: %d a %d (%d cuentas por moneda, %d registros nuevos)"
          % (inicio, fin, args.cantidad, len(nuevos)))
    print("Pesos   nuevos: registros %d a %d" % tuple(pos_nuevas["00"]))
    print("Dólares nuevos: registros %d a %d" % tuple(pos_nuevas["02"]))
    print("Total de detalles: %d" % n_det)
    print("BT: detalles=%s total saldos=%s (%d)"
          % (out_bt[25:32], out_bt[9:25], total_saldos))
    print("FT: detalles=%s" % out_ft[9:16])
    print("Salida: %s (%d bytes)" % (salida, len(cuerpo)))


if __name__ == "__main__":
    main()
