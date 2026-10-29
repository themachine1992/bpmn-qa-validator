#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
================================================================================
 BPMN QA VALIDATOR  ·  v3
================================================================================
Analiza un archivo BPMN 2.0 XML (exportado desde Signavio u otra herramienta de
modelado) y verifica nomenclatura BPMN y lineamientos de modelado de procesos:
tipos de tarea, pasarelas, calles, eventos de inicio/fin, atributos del
diagrama, nombres con verbo en infinitivo, elementos no habilitados, etc.

Las reglas son configurables: los elementos prohibidos, los atributos
obligatorios y las palabras clave de deteccion estan en constantes al inicio
del archivo, para adaptarlas al estandar de modelado de cada organizacion.

Uso:
    python bpmn_qa_validator.py proceso.bpmn
    python bpmn_qa_validator.py proceso.bpmn --salida reporte_qa.xlsx

Diseno:
  - Cada regla tiene ID (Rxx), severidad y apunta a un criterio del estandar.
  - AUTO = validacion completa desde el XML.
  - PARC = detecta indicios; requiere confirmacion humana (se marca en el detalle).
  - INFO = metrica informativa (no es pass/fail).
  - Lo NO automatizable (logica/coherencia) queda para la revision humana.

Reglas: R01 calles sin tareas, R02 tipo de tarea, R03 sistema asociado,
R04 etiquetas de pasarela, R05 flujos convergentes, R06 inicio/fin,
R07 nodos desconectados, R08 verbo infinitivo, R09 mensajes por eventos,
R10 subprocesos, R11-R15 atributos obligatorios, R16 RPA, R17 escalamiento,
R18 calle cliente, R19 portal externo, R20 calles proveedor, R21 minimo de
actividades, R22 pasarela compleja, R23 lista negra de elementos, R24 fin
simple, R25 nombres de calle, R26 frecuencia, R27 balance de pasarelas
paralelas, R28 anotaciones sin asociar, R29 indice de digitalizacion (INFO).
================================================================================
"""

import sys
import os
import re
import argparse
import xml.etree.ElementTree as ET
from collections import defaultdict

# --- Namespaces BPMN 2.0 / DI (los que exportan Signavio y otras herramientas) ----------------------
NS = {
    "bpmn":   "http://www.omg.org/spec/BPMN/20100524/MODEL",
    "bpmndi": "http://www.omg.org/spec/BPMN/20100524/DI",
    "omgdi":  "http://www.omg.org/spec/DD/20100524/DI",
    "omgdc":  "http://www.omg.org/spec/DD/20100524/DC",
}

# --- Catalogos del estandar de modelado ---------------------------------------
# Tags de tarea tipificadas (tipo de tarea)
TASK_TAGS_TIPIFICADAS = {
    "userTask": "Usuario",
    "serviceTask": "Servicio",
    "manualTask": "Manual",
    "businessRuleTask": "Regla de Negocio (RPA)",
}
TASK_TAG_GENERICA = "task"  # <task> sin tipo = incumple "Tipo de Tarea"
TIPOS_QUE_REQUIEREN_SISTEMA = {"userTask", "serviceTask", "businessRuleTask"}

# Tags BPMN prohibidos por la estandar de modelado ("No se encuentra en uso" / no habilitado)
# v3: lista negra COMPLETA segun el estandar de modelado.
TAGS_PROHIBIDOS = {
    # --- Tipos de tarea no habilitados (4.2.1) ---
    "sendTask":     "Tipo de tarea 'Enviar' no habilitado (usar Evento de mensaje).",
    "receiveTask":  "Tipo de tarea 'Recibir' no habilitado.",
    "scriptTask":   "Tipo de tarea 'Script' no habilitado.",
    # --- Pasarelas (4.1.2) ---
    "complexGateway": "Compuerta Compleja no recomendada por su complejidad de lectura.",
    # --- Calles de natacion (4.1.3) 'No se encuentra en uso' ---
    "participantMultiplicityMarker": "Participante Adicional no habilitado.",
    # --- Objetos de datos (4.1.4) 'No se encuentra en uso' ---
    "message":      "Objeto 'Mensaje' no habilitado (usar Evento de mensaje).",
    # --- Artefactos (4.1.5) 'No se encuentra en uso' ---
    "group":        "Artefacto 'Grupo' no habilitado.",
}

# Subtipos que se detectan por eventDefinition dentro de eventos (4.1.6 / 4.1.7)
# Estos NO son tags propios sino definiciones internas de startEvent/endEvent.
EVENT_DEFS_PROHIBIDOS = {
    # Eventos de inicio en desuso (4.1.6)
    "errorEventDefinition_start": "Evento de inicio erroneo no habilitado.",
    "compensateEventDefinition_start": "Evento de inicio de compensacion no habilitado.",
    "signalEventDefinition_start": "Evento de senal de inicio no habilitado.",
}

# Subprocesos en desuso (4.1.1): expandido y de evento.
# subProcess con triggeredByEvent=true = subproceso de evento (no habilitado).
# subProcess con isExpanded (no colapsado) = subproceso expandido (no habilitado).

# Atributos personalizados obligatorios del DIAGRAMA (checklist)
# clave -> (etiqueta legible, regla_id, patron_regex opcional para formato)
ATRIBUTOS_OBLIGATORIOS_DIAGRAMA = {
    "codigo_regional":  ("Codigo Regional", "R11", None),
    "unidad_negocio":   ("Unidad de Negocio", "R12", None),
    "process_owner":    ("Dueno del proceso (Process Owner)", "R13", None),
    "criticidad":       ("Criticidad del proceso", "R15", None),
}

# Palabras clave para deteccion heuristica de atributos en extensionElements
KEYS_ATRIBUTOS = {
    "codigo_regional":  ["codigo regional", "codigoregional", "regional code", "cod regional"],
    "unidad_negocio":   ["unidad de negocio", "unidadnegocio", "business unit", "unidad negocio"],
    "process_owner":    ["process owner", "dueno", "dueno del proceso", "processowner", "owner"],
    "criticidad":       ["criticidad", "proceso critico", "criticality", "critico"],
    "sistema":          ["sistema", "system", "aplicacion", "application"],
    "estado_rpa":       ["estado rpa", "estadorpa", "rpa status"],
    "portal_externo":   ["portal externo", "portalexterno", "external portal"],
    "frecuencia":       ["frecuencia", "frequency", "frequency per year", "frequencyperyear"],
    "referencia_subproceso": ["referencia subprocesal", "referencia subproceso", "subprocess reference"],
}

# Colores validos para "Portal Externo" (naranja). Rango aproximado en HEX.
NARANJA_HEX = ["ff9900", "ffa500", "ff8c00", "ffa64d", "ff9933", "e69138", "f6b26b", "ffb366"]

VERBO_INFINITIVO_RE = re.compile(r".+(ar|er|ir)$", re.IGNORECASE)


def local(tag):
    """Nombre local de un tag, sin namespace."""
    return tag.split("}")[-1] if "}" in tag else tag


def local_es(tag):
    mapa = {
        "task": "Actividad", "userTask": "Actividad Usuario",
        "serviceTask": "Actividad Servicio", "manualTask": "Actividad Manual",
        "businessRuleTask": "Actividad RPA", "subProcess": "Subproceso",
        "callActivity": "Actividad de llamada",
        "exclusiveGateway": "Compuerta XOR", "parallelGateway": "Compuerta Paralela",
        "inclusiveGateway": "Compuerta OR", "complexGateway": "Compuerta Compleja",
        "eventBasedGateway": "Compuerta por Evento",
        "startEvent": "Evento Inicio", "endEvent": "Evento Fin",
        "intermediateThrowEvent": "Evento Intermedio (lanz.)",
        "intermediateCatchEvent": "Evento Intermedio (capt.)",
    }
    return mapa.get(tag, tag)


# ============================================================== HALLAZGO
class Hallazgo:
    def __init__(self, regla_id, severidad, tipo_auto, criterio, elemento, nombre, detalle):
        self.regla_id = regla_id
        self.severidad = severidad          # BLOQUEANTE / ALTA / MEDIA / BAJA
        self.tipo_auto = tipo_auto          # AUTO / PARC
        self.criterio = criterio
        self.elemento = elemento
        self.nombre = nombre or "(sin nombre)"
        self.detalle = detalle

    def as_row(self):
        return [self.regla_id, self.severidad, self.tipo_auto, self.criterio,
                self.elemento, self.nombre, self.detalle]


# ============================================================== VALIDADOR
class BpmnQAValidator:
    def __init__(self, path):
        self.path = path
        self.tree = ET.parse(path)
        self.root = self.tree.getroot()
        self.hallazgos = []
        self.flows = []
        self.flow_in = defaultdict(list)
        self.flow_out = defaultdict(list)
        self.nodes = {}
        self.node_lane = {}
        self.lanes = []
        self.di_fill = {}      # bpmnElement id -> color de fondo (HEX sin #)
        self._indexar()

    # -------------------------------------------------------------- indice
    def _indexar(self):
        for elem in self.root.iter():
            tag = local(elem.tag)
            eid = elem.get("id")
            if eid and tag not in ("BPMNShape", "BPMNEdge", "BPMNPlane", "BPMNDiagram"):
                self.nodes[eid] = elem
            if tag == "sequenceFlow":
                f = (elem.get("id"), elem.get("name"),
                     elem.get("sourceRef"), elem.get("targetRef"))
                self.flows.append(f)
                self.flow_out[f[2]].append(f)
                self.flow_in[f[3]].append(f)

        for lane in self.root.iter():
            if local(lane.tag) != "lane":
                continue
            self.lanes.append(lane)
            for ref in lane:
                if local(ref.tag) == "flowNodeRef":
                    self.node_lane[ref.text] = lane.get("name")

        # Colores de fondo desde BPMN-DI (para Portal Externo)
        for shape in self.root.iter():
            if local(shape.tag) != "BPMNShape":
                continue
            be = shape.get("bpmnElement")
            fill = shape.get("background-color") or shape.get("fill")
            # Signavio suele exponer color en atributos style/di
            for k, v in shape.attrib.items():
                if "color" in k.lower() and v and v.startswith("#"):
                    fill = v
            if be and fill:
                self.di_fill[be] = fill.lstrip("#").lower()

    def add(self, *args):
        self.hallazgos.append(Hallazgo(*args))

    # ---------------------------------------- helper: texto de extensionElements
    def _texto_extension(self, elem):
        """Concatena texto y atributos de extensionElements y meta de un nodo."""
        partes = []
        for child in elem.iter():
            t = local(child.tag).lower()
            if "extensionelements" in t or "property" in t or "meta" in t \
               or "attribute" in t or "signavio" in t:
                if child.text:
                    partes.append(child.text)
                for k, v in child.attrib.items():
                    partes.append(str(k))
                    partes.append(str(v))
        # Tambien atributos directos del nodo (Signavio a veces los pone ahi)
        for k, v in elem.attrib.items():
            partes.append(str(k)); partes.append(str(v))
        return " ".join(partes).lower()

    def _tiene_atributo(self, elem, clave):
        """True si el nodo expone (con valor) el atributo corporativo `clave`."""
        texto = self._texto_extension(elem)
        for kw in KEYS_ATRIBUTOS.get(clave, []):
            if kw in texto:
                # Heuristica: la keyword aparece; asumimos presencia del atributo.
                # Verificamos que no este obviamente vacio ("=" seguido de nada).
                return True
        return False

    # ============================================================== REGLAS

    def run(self):
        # --- Estructurales (AUTO) ---
        self.R01_lanes_sin_tareas()
        self.R04_gateways_bifurcacion_sin_etiqueta()
        self.R05_flujos_convergentes_sin_compuerta()
        self.R07_nodos_desconectados()
        # --- Criterios basicos ---
        self.R02_tipo_de_tarea()
        self.R03_sistema_asociado()
        self.R09_mensajes_via_evento()
        self.R10_subprocesos()
        # --- Atributos obligatorios AUTO ---
        self.R11_R15_atributos_obligatorios()
        self.R14_documentacion_proceso()
        # --- Practicas aceptadas PARC ---
        self.R16_rpa_estado_y_tipo()
        self.R17_intercambios_entre_unidades()
        self.R18_autogestion_cliente()
        self.R19_portal_externo()
        self.R20_proveedores()
        # --- estandar de modelado (varias) ---
        self.R06_entradas_salidas_proceso()
        self.R08_nombre_verbo_infinitivo()
        self.R21_minimo_actividades()
        self.R22_compuerta_compleja()
        self.R23_elementos_prohibidos()
        self.R24_fin_simple_recomendado()
        self.R25_carril_con_nombre()
        self.R26_frecuencia_proceso()
        # --- v3: brechas cerradas + metrica ---
        self.R27_balance_compuertas_paralelas()
        self.R28_anotaciones_sin_asociacion()
        self.R29_indice_digitalizacion()
        return self.hallazgos

    # ---- R01 · Actor/Lane sin tareas (B01, AUTO) ----
    def R01_lanes_sin_tareas(self):
        for lane in self.lanes:
            refs = [r.text for r in lane if local(r.tag) == "flowNodeRef"]
            tiene = any(
                (local(self.nodes[r].tag).endswith("Task") or
                 local(self.nodes[r].tag) in ("task", "subProcess", "callActivity"))
                for r in refs if r in self.nodes)
            if not tiene:
                self.add("R01", "ALTA", "AUTO",
                         "Nomenclatura BPMN - Calles de natacion",
                         "Lane (Actor)", lane.get("name"),
                         "El actor/carril no tiene ninguna actividad asociada. "
                         "Todo carril debe contener al menos una tarea o eliminarse.")

    # ---- R02 · Tipo de Tarea (C01, AUTO, 5%) ----
    def R02_tipo_de_tarea(self):
        for eid, elem in self.nodes.items():
            if local(elem.tag) == TASK_TAG_GENERICA:
                self.add("R02", "ALTA", "AUTO",
                         "Criterios Basicos - Tipo de Tarea",
                         "Actividad", elem.get("name"),
                         "La actividad no tiene definido el 'Tipo de Tarea' "
                         "(Manual, Usuario, Servicio o Regla de Negocio/RPA).")

    # ---- R03 · Sistema asociado (C02/H03, PARC, 5%) ----
    def R03_sistema_asociado(self):
        for eid, elem in self.nodes.items():
            tag = local(elem.tag)
            if tag in TIPOS_QUE_REQUIEREN_SISTEMA:
                if not self._tiene_sistema_asociado(eid, elem):
                    self.add("R03", "ALTA", "PARC",
                             "Criterios Basicos - Asociacion de Sistemas",
                             f"Actividad {TASK_TAGS_TIPIFICADAS.get(tag,'')}",
                             elem.get("name"),
                             "Actividad de Usuario/Servicio/RPA sin sistema asociado. "
                             "PARC: verificar ademas que el sistema pertenezca al diccionario.")

    def _tiene_sistema_asociado(self, eid, elem):
        if self._tiene_atributo(elem, "sistema"):
            return True
        # Asociacion a DataStore/DataObject (Sistema IT / Almacen de datos)
        for f in self.flow_out[eid] + self.flow_in[eid]:
            otro = f[3] if f[2] == eid else f[2]
            if otro in self.nodes and "data" in local(self.nodes[otro].tag).lower():
                return True
        # Asociaciones (association) que apunten a un artefacto de sistema
        for a in self.root.iter():
            if local(a.tag) == "association":
                s, t = a.get("sourceRef"), a.get("targetRef")
                if eid in (s, t):
                    return True
        return False

    # ---- R04 · Compuerta divergente sin Si/No (B02, AUTO) ----
    def R04_gateways_bifurcacion_sin_etiqueta(self):
        for eid, elem in self.nodes.items():
            tag = local(elem.tag)
            if tag not in ("exclusiveGateway", "inclusiveGateway"):
                continue
            salidas = self.flow_out[eid]
            if len(salidas) >= 2:
                if not elem.get("name"):
                    self.add("R04", "ALTA", "AUTO",
                             "Nomenclatura BPMN - Pasarelas (XOR/OR)",
                             "Compuerta divergente", elem.get("name"),
                             "La compuerta de decision no tiene titulo/criterio. "
                             "Debe describir la condicion de la decision.")
                sin_rotulo = [f for f in salidas if not f[1]]
                if sin_rotulo:
                    self.add("R04", "ALTA", "AUTO",
                             "Nomenclatura BPMN - Pasarelas (XOR/OR)",
                             "Compuerta divergente", elem.get("name"),
                             f"{len(sin_rotulo)} de {len(salidas)} ramas de salida no "
                             "describen el camino (Si/No/condicion).")

    # ---- R05 · Convergencia sin compuerta de acoplamiento (B03, AUTO) ----
    def R05_flujos_convergentes_sin_compuerta(self):
        GATEWAYS = ("exclusiveGateway", "inclusiveGateway", "parallelGateway",
                    "complexGateway", "eventBasedGateway")
        for eid, elem in self.nodes.items():
            tag = local(elem.tag)
            entradas = self.flow_in[eid]
            if len(entradas) >= 2 and tag not in GATEWAYS and not tag.endswith("Event"):
                self.add("R05", "MEDIA", "AUTO",
                         "Nomenclatura BPMN - Compuertas de acoplamiento",
                         local_es(tag), elem.get("name"),
                         f"El elemento recibe {len(entradas)} flujos convergentes sin usar "
                         "una compuerta de acoplamiento (merge/join).")

    # ---- R06 · Entradas y salidas del proceso (E02, PARC, 20%) ----
    def R06_entradas_salidas_proceso(self):
        for proc in [e for e in self.root.iter() if local(e.tag) == "process"]:
            inicios = [e for e in proc.iter() if local(e.tag) == "startEvent"]
            fines = [e for e in proc.iter() if local(e.tag) == "endEvent"]
            nombre = proc.get("name") or proc.get("id")
            if not inicios:
                self.add("R06", "BLOQUEANTE", "AUTO",
                         "Entradas y Salidas del Proceso",
                         "Proceso", nombre, "El proceso no tiene Evento de Inicio.")
            if not fines:
                self.add("R06", "BLOQUEANTE", "AUTO",
                         "Entradas y Salidas del Proceso",
                         "Proceso", nombre, "El proceso no tiene Evento de Fin.")
            # PARC: el inicio deberia describir el input (nombre no vacio o nota)
            for s in inicios:
                if not s.get("name"):
                    self.add("R06", "MEDIA", "PARC",
                             "Entradas y Salidas del Proceso",
                             "Evento Inicio", s.get("id"),
                             "El evento de inicio no describe el input que gatilla el "
                             "proceso (rotulo/nota vacio).")
            # v3 - PARC: el fin debe tener rotulo que confirme el cierre
            # (positivo / negativo / escalamiento) segun el estandar de modelado.
            for f in fines:
                if not f.get("name"):
                    self.add("R06", "MEDIA", "PARC",
                             "Entradas y Salidas del Proceso",
                             "Evento Fin", f.get("id"),
                             "El evento de fin no tiene rotulo. Debe confirmar el cierre "
                             "positivo, negativo o el escalamiento del flujo.")

    # ---- R07 · Nodos desconectados (B04, AUTO) ----
    def R07_nodos_desconectados(self):
        for eid, elem in self.nodes.items():
            tag = local(elem.tag)
            es_flujo = (tag.endswith("Task") or tag in
                        ("task", "subProcess", "callActivity", "exclusiveGateway",
                         "parallelGateway", "inclusiveGateway", "complexGateway"))
            if es_flujo and not self.flow_in[eid] and not self.flow_out[eid]:
                self.add("R07", "ALTA", "AUTO",
                         "Nomenclatura BPMN - Conexion de objetos",
                         local_es(tag), elem.get("name"),
                         "Elemento sin flujos de secuencia (desconectado del flujo).")

    # ---- R08 · Nombre con verbo en infinitivo (V01, AUTO) ----
    def R08_nombre_verbo_infinitivo(self):
        for eid, elem in self.nodes.items():
            tag = local(elem.tag)
            if tag.endswith("Task") or tag == "task":
                nombre = (elem.get("name") or "").strip()
                if not nombre:
                    self.add("R08", "MEDIA", "AUTO",
                             "Atributos Principales - Nombre",
                             "Actividad", nombre, "La actividad no tiene nombre.")
                elif not VERBO_INFINITIVO_RE.match(nombre.split()[0]):
                    self.add("R08", "BAJA", "AUTO",
                             "Atributos Principales - Nombre",
                             "Actividad", nombre,
                             "El nombre no inicia con verbo en infinitivo "
                             "(ej: 'Generar', 'Validar', 'Enviar').")

    # ---- R09 · Envio de mensajes via evento, no via actividad (C03, PARC, 5%) ----
    def R09_mensajes_via_evento(self):
        patrones = ("enviar correo", "enviar mensaje", "enviar notific", "notificar",
                    "mandar correo", "enviar email", "enviar mail", "envio de correo",
                    "comunicar por correo")
        for eid, elem in self.nodes.items():
            tag = local(elem.tag)
            if tag.endswith("Task") or tag == "task":
                nombre = (elem.get("name") or "").lower()
                if any(p in nombre for p in patrones):
                    self.add("R09", "MEDIA", "PARC",
                             "Criterios Basicos - Intercambio de Mensajes/Correo",
                             "Actividad", elem.get("name"),
                             "El envio de correos/mensajes debe modelarse con un Evento de "
                             "mensaje, no con una actividad. PARC: confirmar si es envio real.")

    # ---- R10 · Subprocesos con referencia + entradas/salidas (C04/E01, PARC) ----
    def R10_subprocesos(self):
        for eid, elem in self.nodes.items():
            tag = local(elem.tag)
            if tag in ("subProcess", "callActivity"):
                # Referencia subprocesal (atributo)
                if not self._tiene_atributo(elem, "referencia_subproceso") \
                   and tag == "callActivity" and not elem.get("calledElement"):
                    self.add("R10", "MEDIA", "PARC",
                             "Uso de Subprocesos (Entradas/Salidas 10%)",
                             "Subproceso", elem.get("name"),
                             "Subproceso sin referencia al flujo que invoca "
                             "(Referencia Subprocesal / calledElement).")
                # Entradas y salidas: debe tener flujo entrante y saliente
                if not self.flow_in[eid] or not self.flow_out[eid]:
                    self.add("R10", "MEDIA", "PARC",
                             "Uso de Subprocesos (Entradas/Salidas 10%)",
                             "Subproceso", elem.get("name"),
                             "Subproceso sin entradas y/o salidas claramente definidas.")

    # ---- R11-R15 · Atributos obligatorios del diagrama (A01-A05, AUTO, 40%) ----
    def R11_R15_atributos_obligatorios(self):
        procesos = [e for e in self.root.iter() if local(e.tag) == "process"]
        definitions = self.root
        for clave, (etiqueta, regla_id, patron) in ATRIBUTOS_OBLIGATORIOS_DIAGRAMA.items():
            presente = False
            # Buscar en process, definitions y en cualquier nodo de nivel diagrama
            for cont in [definitions] + procesos:
                if self._tiene_atributo(cont, clave):
                    presente = True
                    break
            if not presente:
                peso = {"R11": "5%", "R12": "5%", "R13": "10%", "R15": "10%"}.get(regla_id, "")
                self.add(regla_id, "ALTA", "AUTO",
                         f"Atributos Obligatorios ({peso}) - {etiqueta}",
                         "Diagrama/Proceso",
                         procesos[0].get("name") if procesos else "(proceso)",
                         f"Falta el atributo obligatorio '{etiqueta}' en el diagrama.")

    # ---- R14 · Documentacion/descripcion del proceso (A04, AUTO, 10%) ----
    def R14_documentacion_proceso(self):
        for proc in [e for e in self.root.iter() if local(e.tag) == "process"]:
            doc = None
            for child in proc:
                if local(child.tag) == "documentation":
                    doc = (child.text or "").strip()
                    break
            nombre = proc.get("name") or proc.get("id")
            if not doc:
                self.add("R14", "ALTA", "AUTO",
                         "Atributos Obligatorios - Descripcion",
                         "Proceso", nombre,
                         "El proceso no tiene Descripcion/Documentacion (alcance y objetivo).")
            elif len(doc) < 15:
                self.add("R14", "MEDIA", "AUTO",
                         "Atributos Obligatorios - Descripcion",
                         "Proceso", nombre,
                         "La descripcion del proceso es demasiado breve para explicar el alcance.")

    # ---- R16 · RPA: Estado RPA + Tipo 'Regla de Negocio' (P01, PARC, 4%) ----
    def R16_rpa_estado_y_tipo(self):
        for eid, elem in self.nodes.items():
            tag = local(elem.tag)
            nombre = (elem.get("name") or "")
            # businessRuleTask (RPA) debe tener Estado RPA marcado
            if tag == "businessRuleTask":
                if not self._tiene_atributo(elem, "estado_rpa"):
                    self.add("R16", "MEDIA", "PARC",
                             "Practicas Aceptadas - Sistemas y RPA",
                             "Actividad RPA", nombre,
                             "Actividad RPA (Regla de Negocio) sin atributo 'Estado RPA' marcado.")
            # Si el nombre menciona RPA/Blue Prism pero NO es businessRuleTask
            if ("rpa" in nombre.lower() or "blue prism" in nombre.lower()) \
               and tag != "businessRuleTask":
                self.add("R16", "MEDIA", "PARC",
                         "Practicas Aceptadas - Sistemas y RPA",
                         local_es(tag), nombre,
                         "La actividad refiere a un RPA pero su Tipo de Tarea no es "
                         "'Regla de Negocio'.")

    # ---- R17 · Intercambios entre unidades: escalada + 'Empresas Relacionadas' (P02, PARC) ----
    def R17_intercambios_entre_unidades(self):
        # Detectar eventos de escalada
        escaladas = [e for e in self.nodes.values()
                     if any(local(c.tag) == "escalationEventDefinition" for c in e)]
        lanes_nombres = [(l.get("name") or "").lower() for l in self.lanes]
        hay_empresas_rel = any("empresas relacionadas" in n for n in lanes_nombres)
        for e in escaladas:
            # el evento de escalada deberia estar en carril Empresas Relacionadas
            eid = e.get("id")
            lane = (self.node_lane.get(eid) or "").lower()
            if "empresas relacionadas" not in lane:
                self.add("R17", "MEDIA", "PARC",
                         "Practicas Aceptadas - Intercambios entre Unidades",
                         local_es(local(e.tag)), e.get("name"),
                         "Evento de escalada (intercambio entre unidades) fuera del carril "
                         "'Empresas Relacionadas'.")
        if escaladas and not hay_empresas_rel:
            self.add("R17", "MEDIA", "PARC",
                     "Practicas Aceptadas - Intercambios entre Unidades",
                     "Diagrama", "(carriles)",
                     "Hay eventos de escalada pero no existe el carril 'Empresas Relacionadas'.")

    # ---- R18 · Autogestion: carril 'Cliente' con tareas de Usuario (P03, PARC, 4%) ----
    def R18_autogestion_cliente(self):
        for lane in self.lanes:
            nombre_lane = (lane.get("name") or "").lower()
            if "cliente" in nombre_lane:
                refs = [r.text for r in lane if local(r.tag) == "flowNodeRef"]
                for r in refs:
                    if r in self.nodes:
                        tag = local(self.nodes[r].tag)
                        if tag.endswith("Task") or tag == "task":
                            if tag not in ("userTask", "manualTask"):
                                self.add("R18", "MEDIA", "PARC",
                                         "Practicas Aceptadas - Autogestion",
                                         local_es(tag), self.nodes[r].get("name"),
                                         "Actividad del cliente en carril 'Cliente' que no es "
                                         "de tipo Usuario/Manual (interaccion de autogestion).")

    # ---- R19 · Portal Externo: atributo + color naranja (P04, PARC, 4%) ----
    def R19_portal_externo(self):
        for eid, elem in self.nodes.items():
            marcado_atributo = self._tiene_atributo(elem, "portal_externo")
            color = self.di_fill.get(eid, "")
            es_naranja = any(color.startswith(n[:4]) for n in NARANJA_HEX) if color else False
            if marcado_atributo and not es_naranja:
                self.add("R19", "MEDIA", "PARC",
                         "Practicas Aceptadas - Portal Externo",
                         local_es(local(elem.tag)), elem.get("name"),
                         "Marcado como 'Portal Externo' pero sin color de fondo naranja.")
            if es_naranja and not marcado_atributo:
                self.add("R19", "BAJA", "PARC",
                         "Practicas Aceptadas - Portal Externo",
                         local_es(local(elem.tag)), elem.get("name"),
                         "Elemento naranja sin el atributo 'Portal Externo' marcado.")

    # ---- R20 · Proveedores: Evento Intermedio + carril 'Proveedor X' (P05, PARC, 4%) ----
    def R20_proveedores(self):
        for lane in self.lanes:
            nombre_lane = (lane.get("name") or "").lower()
            if nombre_lane.startswith("proveedor"):
                refs = [r.text for r in lane if local(r.tag) == "flowNodeRef"]
                usa_evento = any(
                    r in self.nodes and "IntermediateEvent" in local(self.nodes[r].tag).replace("intermediate", "Intermediate")
                    for r in refs)
                usa_evento = usa_evento or any(
                    r in self.nodes and "intermediate" in local(self.nodes[r].tag).lower()
                    for r in refs)
                if not usa_evento:
                    self.add("R20", "MEDIA", "PARC",
                             "Practicas Aceptadas - Proveedores",
                             "Lane Proveedor", lane.get("name"),
                             "Carril de proveedor sin uso de Evento Intermedio para "
                             "representar su intervencion.")

    # ---- R21 · Minimo 4 actividades (V02, AUTO) ----
    def R21_minimo_actividades(self):
        for proc in [e for e in self.root.iter() if local(e.tag) == "process"]:
            actividades = [e for e in proc.iter()
                           if local(e.tag).endswith("Task") or
                           local(e.tag) in ("task", "subProcess", "callActivity")]
            if 0 < len(actividades) < 4:
                self.add("R21", "BAJA", "AUTO",
                         "Estandar de modelado - Estructura de Procesos",
                         "Proceso", proc.get("name") or proc.get("id"),
                         f"El proceso tiene solo {len(actividades)} actividad(es). "
                         "Se recomienda un minimo de 4 para considerarse un flujo.")

    # ---- R22 · Compuerta Compleja no recomendada (V03, AUTO) ----
    def R22_compuerta_compleja(self):
        for eid, elem in self.nodes.items():
            if local(elem.tag) == "complexGateway":
                self.add("R22", "MEDIA", "AUTO",
                         "Estandar de modelado - Pasarelas",
                         "Compuerta Compleja", elem.get("name"),
                         "Uso de Compuerta Compleja: no recomendada por su dificultad de "
                         "lectura en la documentacion. Reemplazar por XOR/OR/Paralela.")

    # ---- R23 · Elementos / tipos prohibidos (V04/V05, AUTO) ----
    # v3: lista negra COMPLETA (tags, subprocesos en desuso, pool colapsado,
    #     participante adicional y eventos de inicio prohibidos).
    def R23_elementos_prohibidos(self):
        for eid, elem in self.nodes.items():
            tag = local(elem.tag)

            # 1) Tags directamente prohibidos
            if tag in TAGS_PROHIBIDOS:
                self.add("R23", "ALTA", "AUTO",
                         "Estandar de modelado - Elementos no habilitados",
                         local_es(tag), elem.get("name"),
                         TAGS_PROHIBIDOS[tag])

            # 2) Subprocesos en desuso (4.1.1): expandido y de evento
            if tag == "subProcess":
                if (elem.get("triggeredByEvent") or "").lower() == "true":
                    self.add("R23", "ALTA", "AUTO",
                             "Estandar de modelado - Elementos no habilitados",
                             "Subproceso de Evento", elem.get("name"),
                             "Subproceso de Evento no habilitado. Solo se usa el "
                             "Subproceso Colapsado.")
                elif (elem.get("isExpanded") or "").lower() == "true":
                    self.add("R23", "ALTA", "AUTO",
                             "Estandar de modelado - Elementos no habilitados",
                             "Subproceso Expandido", elem.get("name"),
                             "Subproceso Expandido no habilitado. Usar Subproceso "
                             "Colapsado con referencia al flujo.")

            # 3) Pool colapsado / black box (4.1.3): participant sin processRef
            if tag == "participant" and not elem.get("processRef"):
                self.add("R23", "ALTA", "AUTO",
                         "Estandar de modelado - Elementos no habilitados",
                         "Pool Colapsado", elem.get("name"),
                         "Pool Colapsado ('black box') no habilitado.")

            # 4) Eventos de inicio con definiciones prohibidas (4.1.6)
            if tag == "startEvent":
                for c in elem:
                    d = local(c.tag)
                    if d == "errorEventDefinition":
                        self.add("R23", "ALTA", "AUTO",
                                 "Estandar de modelado - Elementos no habilitados",
                                 "Evento inicio erroneo", elem.get("name"),
                                 "Evento de inicio erroneo no habilitado.")
                    elif d == "compensateEventDefinition":
                        self.add("R23", "ALTA", "AUTO",
                                 "Estandar de modelado - Elementos no habilitados",
                                 "Evento inicio compensacion", elem.get("name"),
                                 "Evento de inicio de compensacion no habilitado.")
                    elif d == "signalEventDefinition":
                        self.add("R23", "ALTA", "AUTO",
                                 "Estandar de modelado - Elementos no habilitados",
                                 "Evento senal de inicio", elem.get("name"),
                                 "Evento de senal de inicio no habilitado.")

    # ---- R24 · Fin simple recomendado (V06, AUTO) ----
    def R24_fin_simple_recomendado(self):
        for eid, elem in self.nodes.items():
            if local(elem.tag) == "endEvent":
                # tiene definicion de evento distinta de simple?
                defs = [local(c.tag) for c in elem if "EventDefinition" in local(c.tag)]
                # se permite messageEndEvent / escalationEndEvent para escalamiento
                no_recomendados = [d for d in defs if d not in
                                   ("messageEventDefinition", "escalationEventDefinition")]
                if no_recomendados:
                    self.add("R24", "BAJA", "AUTO",
                             "Estandar de modelado - Inicio y Fin",
                             "Evento Fin", elem.get("name"),
                             f"Evento de fin tipo '{no_recomendados[0]}'. Se recomienda usar "
                             "fin simple salvo excepcion justificada.")

    # ---- R25 · Carril con nombre (V07, PARC) ----
    def R25_carril_con_nombre(self):
        for lane in self.lanes:
            if not (lane.get("name") or "").strip():
                self.add("R25", "MEDIA", "PARC",
                         "Estandar de modelado - Calles de natacion",
                         "Lane", lane.get("id"),
                         "El carril no tiene nombre (debe llevar el cargo/area responsable).")

    # ---- R26 · Frecuencia del proceso en el inicio (V08, PARC) ----
    def R26_frecuencia_proceso(self):
        for proc in [e for e in self.root.iter() if local(e.tag) == "process"]:
            inicios = [e for e in proc.iter() if local(e.tag) == "startEvent"]
            for s in inicios:
                if not self._tiene_atributo(s, "frecuencia"):
                    self.add("R26", "BAJA", "PARC",
                             "Estandar de modelado - Otros Atributos",
                             "Evento Inicio", s.get("name") or s.get("id"),
                             "El evento de inicio no indica la 'Frecuencia del proceso' "
                             "(Frequency per Year).")

    # ---- R27 · Balance de compuertas paralelas (AUTO) ----
    # "Cada vez que se abra uno de estos portales, debe cerrarse con uno de estos mismos".
    def R27_balance_compuertas_paralelas(self):
        for proc in [e for e in self.root.iter() if local(e.tag) == "process"]:
            aperturas, cierres = [], []
            for e in proc.iter():
                if local(e.tag) != "parallelGateway":
                    continue
                eid = e.get("id")
                n_out = len(self.flow_out[eid])
                n_in = len(self.flow_in[eid])
                if n_out >= 2:          # divergente = apertura (fork)
                    aperturas.append(e)
                if n_in >= 2:           # convergente = cierre (join)
                    cierres.append(e)
            if len(aperturas) != len(cierres):
                self.add("R27", "MEDIA", "AUTO",
                         "Estandar de modelado - Pasarelas Paralelas",
                         "Proceso", proc.get("name") or proc.get("id"),
                         f"Desbalance de compuertas paralelas: {len(aperturas)} apertura(s) "
                         f"vs. {len(cierres)} cierre(s). Cada fork paralelo debe cerrarse "
                         "con un join paralelo.")

    # ---- R28 · Anotacion de texto sin asociacion (AUTO) ----
    # "Siempre debe estar asociado a un elemento del flujograma".
    def R28_anotaciones_sin_asociacion(self):
        # Recolectar todas las asociaciones (source/target)
        asoc_ids = set()
        for a in self.root.iter():
            if local(a.tag) == "association":
                if a.get("sourceRef"):
                    asoc_ids.add(a.get("sourceRef"))
                if a.get("targetRef"):
                    asoc_ids.add(a.get("targetRef"))
        for eid, elem in self.nodes.items():
            if local(elem.tag) == "textAnnotation":
                if eid not in asoc_ids:
                    self.add("R28", "BAJA", "AUTO",
                             "Estandar de modelado - Artefactos",
                             "Anotacion de Texto", (elem.findtext(".//") or "").strip()[:40],
                             "Anotacion de texto sin asociacion a un elemento del flujograma.")

    # ---- R29 · Indice de digitalizacion / automatizacion (INFO) ----
    # Metrica informativa (no pass/fail). Formulas del PDF:
    #   %Automatizacion = servicio / total
    #   %Digitalizacion = (servicio + RPA + usuario) / total
    def R29_indice_digitalizacion(self):
        for proc in [e for e in self.root.iter() if local(e.tag) == "process"]:
            servicio = rpa = usuario = manual = generica = 0
            for e in proc.iter():
                t = local(e.tag)
                if t == "serviceTask":
                    servicio += 1
                elif t == "businessRuleTask":
                    rpa += 1
                elif t == "userTask":
                    usuario += 1
                elif t == "manualTask":
                    manual += 1
                elif t == "task":
                    generica += 1
            total = servicio + rpa + usuario + manual + generica
            if total == 0:
                continue
            pct_auto = servicio / total * 100
            pct_dig = (servicio + rpa + usuario) / total * 100
            self.add("R29", "INFO", "INFO",
                     "Estandar de modelado - Indice de Digitalizacion",
                     "Proceso", proc.get("name") or proc.get("id"),
                     f"% Automatizacion = {pct_auto:.0f}%  |  % Digitalizacion = {pct_dig:.0f}%  "
                     f"(servicio={servicio}, RPA={rpa}, usuario={usuario}, "
                     f"manual={manual}, sin tipo={generica}, total={total}).")


# ============================================================== SALIDAS
ORDEN_SEV = {"BLOQUEANTE": 0, "ALTA": 1, "MEDIA": 2, "BAJA": 3, "INFO": 4}


def imprimir_consola(hallazgos, archivo):
    hallazgos = sorted(hallazgos, key=lambda h: (ORDEN_SEV.get(h.severidad, 9), h.regla_id))
    print("\n" + "=" * 80)
    print(f" REPORTE DE QA BPMN v3  ·  {os.path.basename(archivo)}")
    print("=" * 80)
    if not hallazgos:
        print(" OK - Sin hallazgos. El diagrama cumple las reglas automatizadas.")
        return
    resumen = defaultdict(int)
    for h in hallazgos:
        resumen[h.severidad] += 1
    print(" Hallazgos: %d  |  %s" % (
        len(hallazgos), "  ".join(f"{k}: {v}" for k, v in
                                  sorted(resumen.items(), key=lambda x: ORDEN_SEV.get(x[0], 9)))))
    print("-" * 80)
    for h in hallazgos:
        print(f"[{h.severidad:<10}][{h.tipo_auto}] {h.regla_id}  {h.elemento} -> {h.nombre}")
        print(f"            {h.detalle}")
        print(f"            Criterio: {h.criterio}")
    print("=" * 80)


def exportar_excel(hallazgos, salida):
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "QA BPMN"
    VERDE = "1F3A5F"
    headers = ["Regla", "Severidad", "Tipo", "Criterio del estandar",
               "Elemento", "Nombre", "Detalle del hallazgo"]
    ws.merge_cells("A1:G1")
    ws["A1"] = "BPMN QA REPORT · v3"
    ws["A1"].font = Font(bold=True, size=12, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor=VERDE)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.append(headers)
    for c in ws[2]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="2E5C8A")
        c.alignment = Alignment(vertical="center", wrap_text=True, horizontal="center")
    colores = {"BLOQUEANTE": "F4CCCC", "ALTA": "FCE5CD", "MEDIA": "FFF2CC",
               "BAJA": "E2EFDA", "INFO": "DDEBF7"}
    hallazgos = sorted(hallazgos, key=lambda h: (ORDEN_SEV.get(h.severidad, 9), h.regla_id))
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for h in hallazgos:
        ws.append(h.as_row())
        fila = ws[ws.max_row]
        fila[1].fill = PatternFill("solid", fgColor=colores.get(h.severidad, "FFFFFF"))
        for c in fila:
            c.alignment = Alignment(vertical="top", wrap_text=True)
            c.border = border
    anchos = [8, 12, 7, 30, 22, 26, 55]
    for i, w in enumerate(anchos, 1):
        ws.column_dimensions[chr(64 + i)].width = w
    ws.freeze_panes = "A3"
    wb.save(salida)
    print(f"\nReporte Excel generado: {salida}")


# ============================================================== MAIN
def main():
    ap = argparse.ArgumentParser(description="BPMN QA Validator v3")
    ap.add_argument("archivo", nargs="?", help="Ruta al archivo .bpmn (BPMN 2.0 XML)")
    ap.add_argument("--salida", default=None, help="Ruta del reporte .xlsx")
    args = ap.parse_args()

    if not args.archivo:
        ap.error("Debe indicar un archivo .bpmn")

    validador = BpmnQAValidator(args.archivo)
    hallazgos = validador.run()
    imprimir_consola(hallazgos, args.archivo)
    if args.salida:
        exportar_excel(hallazgos, args.salida)


if __name__ == "__main__":
    main()
