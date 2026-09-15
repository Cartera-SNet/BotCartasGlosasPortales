"""Catalogo central de IPS. El NIT es la identidad estable; el nombre solo ayuda a resolver casos antiguos."""
import re
import unicodedata

CATALOGO_INICIAL = {
    "900267064": "INVERSIONES AZALUD CLINICA BAHIA",
    "900827065": "CENTRO DE DIAGNOSTICO E IMAGENES BAHIA",
    "900657731": "CENTRO MEDICO Y DE REHABILITACION BAHIA",
    "900826509": "RED DE URGENCIAS DEL MAGDALENA",
    "900513306": "FUNDACION MARIA REINA",
    "900600550": "INVERSIONES MEDICAS BARU",
    "900954800": "CENTRO MEDICO Y DE REHABILITACION BARU",
    "900631361": "INVERSIONES MEDICAS VALLESALUD",
    "900257333": "ODONTOTRANS",
    "901081281": "URGETRAUMA",
    "900792417": "RED DE URGENCIAS DE LA COSTA PACIFICA",
    "901959993": "CLINICA CORDIALIDAD",
    "900002780": "FUNDACION CAMPBELL",
    "901523868": "MOVID IPS SAS",
    "901057487": "TECNOLOGIA DIAGNOSTICA DEL VALLE",
    "900558595": "FUNDACION MEDICA CAMPBELL",
    "901149757": "UNIDAD MEDICA DE TRAUMA DEL VALLE",
    "900900754": "CLINICA VALLE SALUD SAN FERNANDO",
    "900469882": "CENTRO MEDICO SERVISALUD INTEGRAL IPS SAS",
    "802024329": "RED DE URGENCIAS DE LA COSTA LTDA",
    "900847382": "CENTRO MEDICO Y DE REHABILITACION VALLE SALUD",
    "800255591": "MEDICO QUIRURGICA TULUA",
}

ALIASES_INICIALES = {
    "CLINICA BAHIA": "900267064",
    "INVERSIONES AZALUD - CLINICA BAHIA": "900267064",
    "CDI BAHIA": "900827065",
    "CENTRO DE DIAGNOSTICO E IMAGEN BAHIA": "900827065",
    "CMR BAHIA": "900657731",
    "CMYR BAHIA": "900657731",
    "CLINICA BARU": "900600550",
    "BARU": "900600550",
    "VALLE SALUD NORTE": "900631361",
    "RUC PACIFICA": "900792417",
    "RUC MAGDALENA": "900826509",
    "SAN FERNANDO": "900900754",
    "URGETRAUMA SAN FERNANDO": "901081281",
    "ALVERNIA TULUA": "901081281",
    "CMR VALLESALUD": "900847382",
    "MOVID IPS SAS": "901523868",
}


def normalizar_nit(valor):
    digitos = re.sub(r"\D", "", str(valor or ""))
    return digitos or None


def normalizar_nombre(valor):
    texto = unicodedata.normalize("NFKD", str(valor or ""))
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    texto = texto.upper().replace("_", " ")
    return re.sub(r"[^A-Z0-9]+", " ", texto).strip()


def resolver(nit=None, nombre_detectado=None, catalogo=None, aliases=None):
    """Devuelve identidad estable sin inventar una IPS cuando no hay datos."""
    nit = normalizar_nit(nit)
    catalogo = catalogo or CATALOGO_INICIAL
    aliases = aliases or ALIASES_INICIALES
    if nit and nit in catalogo:
        return {"nit": nit, "nombre_estandar": catalogo[nit], "nombre_detectado": nombre_detectado or catalogo[nit], "metodo": "NIT"}
    nombre_normalizado = normalizar_nombre(nombre_detectado)
    nit_por_nombre = {normalizar_nombre(v): k for k, v in catalogo.items()}
    nit_por_nombre.update({normalizar_nombre(k): v for k, v in aliases.items()})
    nit_resuelto = nit_por_nombre.get(nombre_normalizado)
    if nit_resuelto:
        return {"nit": nit_resuelto, "nombre_estandar": catalogo.get(nit_resuelto, nombre_detectado), "nombre_detectado": nombre_detectado or "", "metodo": "NOMBRE"}
    return {"nit": nit, "nombre_estandar": "IPS_NO_IDENTIFICADA", "nombre_detectado": nombre_detectado or "", "metodo": "NO_IDENTIFICADA"}


def validar_identidad(valor):
    return str(valor or "").strip() or None
