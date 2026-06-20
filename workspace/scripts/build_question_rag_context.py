#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Recupera de Qdrant el bloque RAG exacto por question_id para todas las preguntas del caso.
Genera:
- 03_analysis/question_rag_context.json
- 03_analysis/ai_batches/batch_001.json, batch_002.json...

Uso:
python /data/scripts/build_question_rag_context.py \
  --case-dir "/data/cases/Codigo_del_proyecto" \
  --collection tfm_rag_manual \
  --qdrant-url http://qdrant:6333 \
  --catalog-json /data/rag_manual/catalog_questions.json \
  --batch-size 10
"""

import argparse
import json
import re
from pathlib import Path
from collections import OrderedDict

import requests


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
            return str(q[key]).strip()
    # fallback: try to locate Q_XXX in text
    text = json.dumps(q, ensure_ascii=False)
    m = re.search(r"Q_\d{3}", text)
    return m.group(0) if m else ""


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
        "doc_type", "chunk_id", "question_id", "dimension", "bloque", "apartado", "requisito",
        "question_text", "control_id", "control_ids", "control_text",
        "incumplimiento_identificado", "controles_asociados", "amenazas",
        "risk_id", "risk_name", "risk_description",
        "severity_by_tier", "severity_source",
        "mitigation", "mitigation_guide", "conditional_rules", "raw_observations",
        "report_context_example", "report_detail_example", "traceability", "evaluation"
    ]
    return {k: payload.get(k) for k in keys if k in payload}


def make_units(questions_with_catalog):
    """
    Mantiene juntas las preguntas que se evalúan de forma grupal:
    - Q_010-Q_021 autenticación dependiente de federación.
    - Q_037-Q_040 monitorización/logs/SIEM/alertas.
    El resto se empaqueta individualmente.
    """
    grouped = OrderedDict()
    singles = []

    for item in questions_with_catalog:
        eval_info = (item.get("catalog_entry") or {}).get("evaluation") or {}
        group_id = eval_info.get("group_id") or ""
        if group_id:
            grouped.setdefault(group_id, {
                "unit_type": "group",
                "group_id": group_id,
                "group_reason": eval_info.get("group_reason", ""),
                "questions": []
            })
            grouped[group_id]["questions"].append(item)
        else:
            singles.append({
                "unit_type": "single",
                "group_id": "",
                "group_reason": "",
                "questions": [item]
            })

    return list(grouped.values()) + singles


def make_batches(units, batch_size):
    batches = []
    current = []
    count = 0

    for unit in units:
        unit_count = len(unit["questions"])
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
    ap.add_argument("--batch-size", type=int, default=10)
    args = ap.parse_args()

    case_dir = Path(args.case_dir)
    jsondir = case_dir / "02_json"
    analysis_dir = case_dir / "03_analysis"
    batches_dir = analysis_dir / "ai_batches"
    batches_dir.mkdir(parents=True, exist_ok=True)

    triaje = read_json(jsondir / "triaje.json")
    questionnaire = read_json(jsondir / "questionnaire.json")
    case_meta = read_json(jsondir / "case_meta.json")

    local_catalog = load_local_catalog(args.catalog_json)

    questions = questionnaire.get("questions", [])
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

    units = make_units(questions_with_catalog)
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
        "questions_total": len(questions),
        "questions_with_catalog_found": sum(1 for q in questions_with_catalog if str(q["retrieval"]["status"]).startswith("FOUND")),
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
        "questions_with_catalog_found": context["questions_with_catalog_found"]
    }, ensure_ascii=False))


if __name__ == "__main__":
    main()
