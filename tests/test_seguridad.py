# -*- coding: utf-8 -*-
"""
Pruebas de las protecciones del modo --web.

Se lanzan con:   python -m pytest
No necesitan Internet: las fotos las sirve un pequeño servidor web que se
arranca dentro de las propias pruebas.
"""
import io
import os
import sys
import threading
from http.server import BaseHTTPRequestHandler, HTTPServer

import openpyxl
import pytest
from PIL import Image

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
import etiquetar_colores_local as et  # noqa: E402


# ---------------------------------------------------------------
# Servidor de pruebas: /foto.jpg devuelve una foto pequeña y
# /redirige manda a una direccion interna (10.0.0.1).
# ---------------------------------------------------------------
def _jpg(lado=50):
    buf = io.BytesIO()
    Image.new("RGB", (lado, lado), (200, 20, 30)).save(buf, "JPEG")
    return buf.getvalue()


class _Manejador(BaseHTTPRequestHandler):
    def do_GET(self):
        if self.path == "/redirige":
            self.send_response(302)
            self.send_header("Location", "http://10.0.0.1/foto.jpg")
            self.end_headers()
            return
        datos = _jpg(400 if self.path == "/grande.jpg" else 50)
        self.send_response(200)
        self.send_header("Content-Type", "image/jpeg")
        self.end_headers()
        self.wfile.write(datos)

    def log_message(self, *args):
        pass   # que no ensucie la salida de las pruebas


@pytest.fixture(scope="module")
def servidor():
    srv = HTTPServer(("127.0.0.1", 0), _Manejador)
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    yield "http://127.0.0.1:%d" % srv.server_port
    srv.shutdown()


@pytest.fixture(autouse=True)
def limites_normales():
    """Cada prueba empieza y termina con los limites de escritorio."""
    et.configurar_limites(False)
    yield
    et.configurar_limites(False)


# ---------------------------------------------------------------
# Direcciones internas
# ---------------------------------------------------------------
@pytest.mark.parametrize("ip", ["8.8.8.8", "1.1.1.1", "2606:4700:4700::1111"])
def test_ip_publica_se_acepta(ip):
    assert et.ip_es_publica(ip)


@pytest.mark.parametrize("ip", [
    "127.0.0.1",          # el propio servidor
    "10.0.0.5",           # red privada
    "172.16.0.1",         # red privada
    "192.168.1.10",       # red de casa
    "169.254.169.254",    # datos privados del proveedor de hosting
    "0.0.0.0",
    "::1",                # el propio servidor en IPv6
    "fe80::1",            # red local en IPv6
    "::ffff:127.0.0.1",   # IPv4 interna disfrazada de IPv6
])
def test_ip_interna_se_rechaza(ip):
    assert not et.ip_es_publica(ip)


@pytest.mark.parametrize("url", [
    "http://127.0.0.1/foto.jpg",
    "http://localhost/foto.jpg",
    "http://192.168.1.1/",
    "http://[::1]/foto.jpg",
])
def test_web_bloquea_direcciones_internas(url):
    et.configurar_limites(True)
    with pytest.raises(et.DescargaRechazada):
        et.bajar_bytes(url)


def test_web_bloquea_puertos_raros():
    et.configurar_limites(True)
    with pytest.raises(et.DescargaRechazada, match="puerto"):
        et.bajar_bytes("http://example.com:9000/foto.jpg")


@pytest.mark.parametrize("url", ["file:///C:/Windows/win.ini",
                                 "ftp://example.com/foto.jpg"])
def test_solo_http_y_https(url):
    with pytest.raises(et.DescargaRechazada):
        et.bajar_bytes(url)


def test_web_bloquea_redireccion_a_direccion_interna(servidor, monkeypatch):
    # Para poder probarlo sin Internet, el servidor de pruebas (127.0.0.1)
    # hace de "web publica": se permite su IP y su puerto. Lo que se
    # comprueba es que su redireccion a 10.0.0.1 SI se corta.
    et.configurar_limites(True)
    monkeypatch.setattr(et, "ip_es_publica", lambda ip: ip == "127.0.0.1")
    puerto = int(servidor.rsplit(":", 1)[1])
    monkeypatch.setattr(et, "PUERTOS_WEB", (80, 443, puerto))
    assert et.bajar_bytes(servidor + "/foto.jpg")       # la normal, pasa
    with pytest.raises(et.DescargaRechazada, match="interna"):
        et.bajar_bytes(servidor + "/redirige")


def test_escritorio_sigue_bajando_de_la_red_local(servidor):
    # Sin --web no se bloquea nada: en tu ordenador puedes querer bajar
    # fotos de un servidor de tu empresa.
    assert len(et.bajar_bytes(servidor + "/foto.jpg")) > 100


# ---------------------------------------------------------------
# Tamaño de las fotos
# ---------------------------------------------------------------
def test_foto_demasiado_pesada(servidor, monkeypatch):
    monkeypatch.setitem(et.LIMITES, "max_bytes", 1000)
    with pytest.raises(et.DescargaRechazada, match="pesa"):
        et.bajar_bytes(servidor + "/grande.jpg")


def test_bomba_de_pixeles():
    # 6000 x 6000 = 36 millones de pixeles, pero de un solo color: el PNG
    # ocupa muy poco. Se tiene que rechazar sin llegar a descomprimirlo.
    buf = io.BytesIO()
    Image.new("RGB", (6000, 6000), (0, 0, 0)).save(buf, "PNG")
    assert len(buf.getvalue()) < 500_000
    et.configurar_limites(True)
    with pytest.raises(et.DescargaRechazada, match="pixeles"):
        et.abrir_imagen(buf.getvalue())


# ---------------------------------------------------------------
# Excel: inyeccion de formulas
# ---------------------------------------------------------------
def test_formulas_quedan_como_texto(tmp_path):
    libro = openpyxl.Workbook()
    hoja = libro.active
    hoja["A1"] = '=HYPERLINK("http://malo.example","pulsa")'
    hoja["A2"] = "normal"
    et.neutralizar_formulas(hoja)
    ruta = tmp_path / "salida.xlsx"
    libro.save(ruta)

    hoja = openpyxl.load_workbook(ruta).active
    assert hoja["A1"].data_type == "s"           # texto, no formula
    assert hoja["A1"].value.startswith("=HYPERLINK")
    assert hoja["A2"].value == "normal"


# ---------------------------------------------------------------
# Nombres para las sentencias SQL
# ---------------------------------------------------------------
@pytest.mark.parametrize("nombre", ["productos", "colour", "id_marca",
                                    "tienda.productos", "_tmp2"])
def test_nombre_sql_valido(nombre):
    assert et.nombre_sql_valido(nombre)


@pytest.mark.parametrize("nombre", ["productos; drop table productos",
                                    "mi tabla", "1abc", "", "a.b.c",
                                    "color'--"])
def test_nombre_sql_no_valido(nombre):
    assert not et.nombre_sql_valido(nombre)


# ---------------------------------------------------------------
# Lectura del listado
# ---------------------------------------------------------------
def _csv(tmp_path, filas, cabecera="referencia,marca,imagen"):
    ruta = tmp_path / "listado.csv"
    lineas = [cabecera] + ["REF%03d,7,http://x.example/%d.jpg" % (i, i)
                           for i in range(filas)]
    ruta.write_text("\n".join(lineas), encoding="utf-8")
    return str(ruta)


def test_maximo_de_filas(tmp_path):
    with pytest.raises(ValueError, match="maximo"):
        et.leer_listado(_csv(tmp_path, 12), max_filas=10)


def test_web_solo_acepta_csv(tmp_path):
    ruta = tmp_path / "listado.xlsx"
    ruta.write_bytes(b"no importa")
    with pytest.raises(ValueError, match="csv"):
        et.leer_listado(str(ruta), solo_csv=True)


def test_columna_proveedor_con_otro_nombre(tmp_path):
    ruta = _csv(tmp_path, 3)
    # 'marca' no es un nombre conocido: sin la opcion no se detecta...
    assert et.leer_listado(ruta)[3] is None
    # ...y con --columna-proveedor si
    assert et.leer_listado(ruta, col_proveedor="Marca")[3] == "marca"
