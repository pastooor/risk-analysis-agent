#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Indexa catalog_questions.json en Qdrant de forma controlada.

Uso dentro del contenedor n8n:
python /data/scripts/index_manual_qdrant.py \
  --catalog-json /data/rag_manual/catalog_questions.json \
  --collection tfm_rag_manual \
  --qdrant-url http://qdrant:6333 \
  --ollama-url http://ollama:11434 \
  --embedding-model nomic-embed-text \
  --recreate
"""

import argparse
import json
import uuid
from pathlib import Path

import requests


def load_catalog(path: str):
    data = json.loads(Path(path).read_text(encoding="utf-8"))
    if not isinstance(data, list):
        raise ValueError("catalog_questions.json debe ser una lista de puntos")
    return data


def ollama_embed(ollama_url: str, model: str, text: str):
    url = ollama_url.rstrip("/") + "/api/embeddings"
    r = requests.post(url, json={"model": model, "prompt": text}, timeout=180)
    r.raise_for_status()
    data = r.json()
    emb = data.get("embedding")
    if not isinstance(emb, list) or not emb:
        raise RuntimeError(f"Ollama no devolvió embedding válido: {data}")
    return emb


def qdrant_request(method: str, qdrant_url: str, path: str, **kwargs):
    url = qdrant_url.rstrip("/") + path
    r = requests.request(method, url, timeout=180, **kwargs)
    if r.status_code >= 400:
        raise RuntimeError(f"Qdrant error {r.status_code} {method} {url}: {r.text}")
    if r.text:
        return r.json()
    return {}


def create_collection(qdrant_url: str, collection: str, vector_size: int, recreate: bool):
    if recreate:
        try:
            requests.delete(qdrant_url.rstrip() + f"/collections/{collection}", timeout=60)
        except Exception:
            pass

    body = {
        "vectors": {
            "size": vector_size,
            "distance": "Cosine"
        }
    }
    qdrant_request("PUT", qdrant_url, f"/collections/{collection}", json=body)

    # Índices de payload para búsquedas exactas. Si ya existen, Qdrant puede responder error; lo ignoramos.
    for field in ["doc_type", "question_id", "chunk_id", "evaluation.group_id", "evaluation.evaluation_mode"]:
        try:
            qdrant_request(
                "PUT",
                qdrant_url,
                f"/collections/{collection}/index",
                json={"field_name": field, "field_schema": "keyword"},
            )
        except Exception:
            pass


def point_id_for(record):
    # Qdrant REST acepta IDs enteros o UUID. No acepta strings arbitrarios tipo "RAG-Q_020".
    # Por eso mantenemos el ID lógico en payload.chunk_id y usamos UUID determinista como point id.
    raw = record.get("qdrant_point_id") or record.get("id") or record.get("payload", {}).get("chunk_id")
    try:
        uuid.UUID(str(raw))
        return str(raw)
    except Exception:
        return str(uuid.uuid5(uuid.NAMESPACE_URL, str(raw)))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--catalog-json", required=True)
    ap.add_argument("--collection", default="tfm_rag_manual")
    ap.add_argument("--qdrant-url", default="http://qdrant:6333")
    ap.add_argument("--ollama-url", default="http://ollama:11434")
    ap.add_argument("--embedding-model", default="nomic-embed-text")
    ap.add_argument("--recreate", action="store_true")
    ap.add_argument("--batch-size", type=int, default=32)
    args = ap.parse_args()

    catalog = load_catalog(args.catalog_json)
    if not catalog:
        raise ValueError("Catálogo vacío")

    first_vector = ollama_embed(args.ollama_url, args.embedding_model, catalog[0]["vector_text"])
    vector_size = len(first_vector)

    create_collection(args.qdrant_url, args.collection, vector_size, args.recreate)

    points = []
    for idx, record in enumerate(catalog):
        vector = first_vector if idx == 0 else ollama_embed(args.ollama_url, args.embedding_model, record["vector_text"])
        payload = dict(record.get("payload", {}))
        payload["logical_id"] = record.get("id")
        payload["vector_text"] = record.get("vector_text", "")
        points.append({
            "id": point_id_for(record),
            "vector": vector,
            "payload": payload
        })

        if len(points) >= args.batch_size:
            qdrant_request("PUT", args.qdrant_url, f"/collections/{args.collection}/points?wait=true", json={"points": points})
            points = []

    if points:
        qdrant_request("PUT", args.qdrant_url, f"/collections/{args.collection}/points?wait=true", json={"points": points})

    print(json.dumps({
        "ok": True,
        "collection": args.collection,
        "points_indexed": len(catalog),
        "vector_size": vector_size,
        "embedding_model": args.embedding_model
    }, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
