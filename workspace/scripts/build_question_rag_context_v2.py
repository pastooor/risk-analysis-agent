#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Build deterministic RAG context for SaaS risk analysis.

V2 changes:
- Default batch size reduced to 3.
- AUTHENTICATION Q_010-Q_021 are NOT packed as one huge group; each question is evaluated individually with shared auth_context.
- LOGGING Q_037-Q_040 remains grouped because the catalog requires joint evaluation.
- Adds auth_context and logging_context to every batch/unit so the AI can score with context but cannot mix catalog fields.

Usage:
python /data/scripts/build_question_rag_context_v2.py \
  --case-dir "/data/cases/Codigo_del_proyecto" \
  --collection tfm_rag_manual \
  --qdrant-url http://qdrant:6333 \
  --catalog-json /data/rag_manual/catalog_questions_v2.json \
  --batch-size 3
"""

import argparse
import json
import re
from pathlib import Path
from collections import OrderedDict

import requests

AUTH_QIDS = {f"Q_{i:03d}" for i in range(10, 22)}
LOGGING_QIDS = {"Q_037", "Q_038", "Q_039", "Q_040"}


def read_json(path):
    return json.loads(Path(path).read_text(encoding="utf-8"))


def write_json(path, data):
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    Path(path).write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def qdrant_exact_question(qdrant_url, collection, question_id):
    body = {
        "filter": {
            "must": [
                {"key": "doc_type", "match": {"value": "risk_catalog_question"}},
                {"key": "question_id", "match": {"value": question_id}}
            ]
        },
        "limit": 1,
        "with_payload": True,
        "with_vector": False
    }
    url = qdrant_url.rstrip("/") + f"/collections/{collection}/points/scroll"
    r = requests.post(url, json=body, timeout=90)
    if r.status_code >= 400:
        raise RuntimeError(f"Qdrant error {r.status_code}: {r.text}")
    points = r.json().get("result", {}).get("points", [])
    if not points:
        return None
    return points[0].get("payload")


def load_local_catalog(catalog_json):
    if not catalog_json:
        return {}
    p = Path(catalog_json)
    if not p.exists():
        return {}
    data = read_json(p)
    by_qid = {}
    for record in data:
        payload = record.get("payload", {})
        qid = payload.get("question_id")
        if qid:
            by_qid[qid] = payload
    return by_qid


def get_question_id(q):
    for key in ["id", "question_id", "code", "qid"]:
        if q.get(key):
            return str(q[key]).strip().upper()
    text = json.dumps(q, ensure_ascii=False)
    m = re.search(r"Q_\d{3}", text, re.IGNORECASE)
    return m.group(0).upper() if m else ""


def compact_question(q):
    return {
        "question_id": get_question_id(q),
        "question_text": q.get("question") or q.get("text") or q.get("title") or q.get("pregunta") or "",
        "provider_answer": q.get("answer"),
        "provider_explanation": q.get("explanation") or q.get("details") or q.get("comment") or q.get("comments") or q.get("observations") or "",
        "answer_missing": bool(q.get("answer_missing")),
        "raw_question": q
    }


def compact_catalog(payload):
    if not payload:
        return None
    keys = [
        "doc_type", "chunk_id", "question_id", "question_number",
        "dimension", "bloque", "apartado", "requisito",
        "question_text", "control_id", "control_ids", "control_text",
        "incumplimiento_identificado", "controles_asociados", "amenazas",
        "risk_id", "risk_name", "risk_description",
        "severity_by_tier", "severity_source",
        "evaluation_group", "severity_logic", "depends_on_questions", "conditional_severity_logic",
        "mitigation", "mitigation_guide", "conditional_rules", "raw_observations",
        "report_context_example", "report_detail_example", "traceability", "evaluation"
    ]
    return {k: payload.get(k) for k in keys if k in payload}


def normalize_text(value):
    if value is None:
        return ""
    value = str(value)
    repl = str.maketrans("áéíóúÁÉÍÓÚñÑ", "aeiouAEIOUnN")
    return value.translate(repl).lower().strip()


def find_value_recursive(obj, wanted_keys=None, wanted_labels=None):
    """Best-effort extraction from flexible parsed triaje/questionnaire JSON."""
    wanted_keys = [normalize_text(k) for k in (wanted_keys or [])]
    wanted_labels = [normalize_text(k) for k in (wanted_labels or [])]

    def walk(x):
        if isinstance(x, dict):
            # direct key match
            for k, v in x.items():
                nk = normalize_text(k)
                if nk in wanted_keys and v not in [None, ""]:
                    return v

            label = normalize_text(x.get("label") or x.get("name") or x.get("title") or x.get("question") or "")
            if label and any(w in label for w in wanted_labels):
                for val_key in ["value", "answer", "text", "description"]:
                    if x.get(val_key) not in [None, ""]:
                        return x.get(val_key)

            for v in x.values():
                found = walk(v)
                if found not in [None, ""]:
                    return found
        elif isinstance(x, list):
            for item in x:
                found = walk(item)
                if found not in [None, ""]:
                    return found
        return None

    return walk(obj)


def build_question_map(questions):
    out = {}
    for q in questions:
        cq = compact_question(q)
        if cq["question_id"]:
            out[cq["question_id"]] = cq
    return out


def build_auth_context(questions_by_id, triaje, case_meta):
    q009 = questions_by_id.get("Q_009", {})
    q010 = questions_by_id.get("Q_010", {})
    info_class = (
        find_value_recursive(triaje, ["clasificacion_informacion", "tipo_informacion", "tipo de informacion"], ["tipo de informacion", "clasificacion"])
        or find_value_recursive(case_meta, ["clasificacion_informacion", "tipo_informacion"], ["clasificacion"])
        or ""
    )
    return {
        "description": "Contexto común para Q_010-Q_021. La severidad se resuelve con las observaciones del catálogo, no con TIER directo si severity_source=conditional_rules.",
        "empresa_users_access_question_id": "Q_009",
        "empresa_users_access_answer": q009.get("provider_answer"),
        "empresa_users_access_explanation": q009.get("provider_explanation", ""),
        "federation_question_id": "Q_010",
        "federation_supported_or_used_answer": q010.get("provider_answer"),
        "federation_supported_or_used_explanation": q010.get("provider_explanation", ""),
        "information_classification": info_class,
        "tier": case_meta.get("tier"),
        "rules_summary": [
            "Si no hay accesos de usuarios de EMPRESA, normalmente no se levanta riesgo de autenticación.",
            "Si hay accesos de usuarios de EMPRESA y existe federación corporativa suficiente, los controles alternativos de repositorio, contraseñas, MFA y timeout pueden quedar sin riesgo según catalog_entry.",
            "Si hay accesos de usuarios de EMPRESA, no hay federación y el control concreto no se cumple, se genera hallazgo aplicando las observaciones del catalog_entry.",
            "Q_020 MFA: si no hay federación y no hay MFA, MEDIO si información pública y ALTO para Uso Interno o superior.",
            "Q_011-Q_019: si no hay federación y no cumplen, MEDIO para Uso Interno y ALTO para Confidencial o superior.",
            "Q_021 timeout: si no hay federación y no cumple, MEDIO."
        ]
    }


def build_logging_context(questions_by_id, case_meta):
    return {
        "description": "Contexto común para Q_037-Q_040. Se evalúan conjuntamente: logs, SIEM propio, integración con SIEM de EMPRESA y alertas.",
        "tier": case_meta.get("tier"),
        "questions": {qid: questions_by_id.get(qid, {}) for qid in sorted(LOGGING_QIDS)},
        "rules_summary": [
            "No disponer de SIEM propio no genera riesgo por sí solo si el proveedor permite enviar logs al SIEM de EMPRESA o existe mecanismo equivalente suficiente.",
            "El objetivo es que haya gestión efectiva de logs y alertas de eventos.",
            "Si no se recogen logs, sí puede existir hallazgo.",
            "Si la información sobre logs/SIEM/alertas es insuficiente, score=2 y follow-up."
        ]
    }


def make_units(questions_with_catalog, questions_by_id, auth_context, logging_context):
    """
    V2 packing strategy:
    - Q_010-Q_021 authentication: individual unit + shared auth_context to avoid field mixing.
    - Q_037-Q_040 logging: one grouped unit because catalog explicitly says joint evaluation.
    - Everything else: individual unit.
    """
    units = []
    logging_group = []

    for item in questions_with_catalog:
        qid = item.get("question_id")
        if qid in LOGGING_QIDS:
            logging_group.append(item)
            continue

        if qid in AUTH_QIDS:
            units.append({
                "unit_type": "single_with_group_context",
                "group_id": "AUTHENTICATION_FEDERATION_DEPENDENCIES",
                "group_reason": "Pregunta de autenticación evaluada individualmente, pero con contexto común de acceso de usuarios de EMPRESA y federación.",
                "shared_context": auth_context,
                "questions": [item]
            })
            continue

        units.append({
            "unit_type": "single",
            "group_id": "",
            "group_reason": "",
            "shared_context": {},
            "questions": [item]
        })

    if logging_group:
        logging_group = sorted(logging_group, key=lambda x: x.get("question_id", ""))
        units.append({
            "unit_type": "group",
            "group_id": "LOGGING_SIEM_ALERTING",
            "group_reason": "Las preguntas Q_037-Q_040 se evalúan conjuntamente según observaciones del catálogo.",
            "shared_context": logging_context,
            "questions": logging_group
        })

    return units


def make_batches(units, batch_size):
    batches = []
    current = []
    count = 0

    for unit in units:
        unit_count = len(unit.get("questions", []))
        # grouped units are atomic even if unit_count > batch_size
        if current and count + unit_count > batch_size:
            batches.append(current)
            current = []
            count = 0
        current.append(unit)
        count += unit_count

    if current:
        batches.append(current)

    return batches


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--case-dir", required=True)
    ap.add_argument("--collection", default="tfm_rag_manual")
    ap.add_argument("--qdrant-url", default="http://qdrant:6333")
    ap.add_argument("--catalog-json", default="")
    ap.add_argument("--batch-size", type=int, default=3)
    args = ap.parse_args()

    case_dir = Path(args.case_dir)
    jsondir = case_dir / "02_json"
    analysis_dir = case_dir / "03_analysis"
    batches_dir = analysis_dir / "ai_batches"
    batches_dir.mkdir(parents=True, exist_ok=True)

    # Clean previous batch files so old batch_010 etc. do not remain
    for old in batches_dir.glob("batch_*.json"):
        old.unlink()

    triaje = read_json(jsondir / "triaje.json")
    questionnaire = read_json(jsondir / "questionnaire.json")
    case_meta = read_json(jsondir / "case_meta.json")

    local_catalog = load_local_catalog(args.catalog_json)

    questions = questionnaire.get("questions", [])
    questions_by_id = build_question_map(questions)
    auth_context = build_auth_context(questions_by_id, triaje, case_meta)
    logging_context = build_logging_context(questions_by_id, case_meta)

    questions_with_catalog = []
    for q in questions:
        cq = compact_question(q)
        qid = cq["question_id"]

        catalog_entry = None
        retrieval_status = "NOT_FOUND"
        retrieval_method = "qdrant_exact_question_id"

        if qid:
            try:
                catalog_entry = qdrant_exact_question(args.qdrant_url, args.collection, qid)
                if catalog_entry:
                    retrieval_status = "FOUND"
            except Exception as e:
                retrieval_status = f"QDRANT_ERROR: {e}"

        if not catalog_entry and qid in local_catalog:
            catalog_entry = local_catalog[qid]
            retrieval_status = "FOUND_LOCAL_FALLBACK"
            retrieval_method = "local_catalog_question_id"

        questions_with_catalog.append({
            **cq,
            "catalog_entry": compact_catalog(catalog_entry),
            "retrieval": {
                "status": retrieval_status,
                "method": retrieval_method,
                "collection": args.collection
            }
        })

    units = make_units(questions_with_catalog, questions_by_id, auth_context, logging_context)
    packed_batches = make_batches(units, args.batch_size)

    batch_refs = []
    for i, units_batch in enumerate(packed_batches, start=1):
        batch_questions = []
        for unit in units_batch:
            batch_questions.extend(unit["questions"])

        batch = {
            "batch_id": f"batch_{i:03d}",
            "case_id": case_dir.name,
            "case_meta": case_meta,
            "auth_context": auth_context,
            "logging_context": logging_context,
            "triaje_summary": {
                "contract": triaje.get("contract", {}),
                "solution": triaje.get("solution", {}),
                "information": triaje.get("information", {}),
                "provider_interaction": triaje.get("provider_interaction", {}),
                "risk_context": triaje.get("risk_context", {}),
                "derived_flags": triaje.get("derived_flags", {})
            },
            "questionnaire_provider": questionnaire.get("provider", {}),
            "units": units_batch,
            "questions": batch_questions
        }
        batch_path = batches_dir / f"batch_{i:03d}.json"
        write_json(batch_path, batch)
        batch_refs.append({
            "batch_id": batch["batch_id"],
            "path": str(batch_path),
            "question_count": len(batch_questions),
            "unit_count": len(units_batch)
        })

    context = {
        "ok": True,
        "case_id": case_dir.name,
        "collection": args.collection,
        "retrieval_strategy": "exact_payload_filter_by_question_id",
        "batching_strategy": "v2_small_batches_auth_individual_with_context_logging_grouped",
        "questions_total": len(questions),
        "questions_with_catalog_found": sum(1 for q in questions_with_catalog if str(q["retrieval"]["status"]).startswith("FOUND")),
        "auth_context": auth_context,
        "logging_context": logging_context,
        "questions_with_catalog": questions_with_catalog,
        "batches": batch_refs,
        "analysis_dir": str(analysis_dir),
        "batches_dir": str(batches_dir)
    }

    write_json(analysis_dir / "question_rag_context.json", context)
    print(json.dumps({
        "ok": True,
        "case_id": case_dir.name,
        "question_rag_context_file": str(analysis_dir / "question_rag_context.json"),
        "batches": batch_refs,
        "questions_total": context["questions_total"],
        "questions_with_catalog_found": context["questions_with_catalog_found"],
        "batching_strategy": context["batching_strategy"]
    }, ensure_ascii=False))


if __name__ == "__main__":
    main()
