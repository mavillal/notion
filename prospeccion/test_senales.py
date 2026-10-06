from senales import analizar_pagina, emails_genericos, links_interesantes


def test_solo_emails_genericos():
    texto = "info@empresa.cl juan.perez@empresa.cl Ventas@Empresa.cl."
    assert emails_genericos(texto) == ["info@empresa.cl", "ventas@empresa.cl"]


def test_detecta_senales():
    html = """<html><body>Certificación ISO 45001 y protocolo Ley Karin.
    Trabaja con nosotros. <a href="mailto:contacto@empresa.cl">x</a></body></html>"""
    r = analizar_pagina(html, "https://empresa.cl")
    assert {"sst_iso45001", "sst_ley_karin", "crecimiento_empleo"} <= set(r["senales"])
    assert r["emails"] == ["contacto@empresa.cl"]


def test_links_mismo_dominio():
    html = '<a href="/contacto">Contacto</a><a href="https://otro.cl/contacto">x</a><a href="/blog">Blog</a>'
    assert links_interesantes(html, "https://www.empresa.cl/", 5) == ["https://www.empresa.cl/contacto"]
